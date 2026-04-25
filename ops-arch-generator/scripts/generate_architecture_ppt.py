#!/usr/bin/env python3
"""Generate a customer-specific PowerPoint architecture diagram from an ops workbook."""

from __future__ import annotations

import argparse
import base64
import json
import math
import re
import sys
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from upload_artifact import load_upload_config, upload_file  # noqa: E402


NS_XLS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/package/2006/relationships",
}

BASE_SLIDE_W = 12192000
BASE_SLIDE_H = 6858000
ASSET_DIR = SCRIPT_DIR.parent / "assets"


def canonical_text(value: str | None) -> str:
    if not value:
        return ""
    value = str(value).strip().lower()
    value = value.replace("（", "(").replace("）", ")").replace("：", ":")
    value = re.sub(r"[\s_:/|,;，；]+", "", value)
    return value


def split_nonempty_lines(value: str | None) -> list[str]:
    if not value:
        return []
    return [line.strip() for line in re.split(r"[\r\n]+", str(value)) if line.strip()]


def uniq_preserve(items: Iterable) -> list:
    seen = set()
    result = []
    for item in items:
        key = json.dumps(item, sort_keys=True, ensure_ascii=False) if isinstance(item, dict) else item
        if key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result


def extract_ip_like(value: str | None) -> str | None:
    if not value:
        return None
    match = re.search(r"((?:\d+|XX)\.(?:\d+|XX)\.(?:\d+|XX)\.(?:\d+))", normalize_endpoint_text(value), re.IGNORECASE)
    return match.group(1) if match else None


def extract_last_octet(value: str | None) -> str | None:
    if not value:
        return None
    match = re.search(r"(?:\d+|XX)\.(?:\d+|XX)\.(?:\d+|XX)\.(\d+)", normalize_endpoint_text(value), re.IGNORECASE)
    return match.group(1) if match else None


def extract_ports_freeform(value: str | None) -> list[str]:
    if not value:
        return []
    return uniq_preserve(re.findall(r"(?<![\d.])(\d{2,5})(?![\d.])", str(value)))


def normalize_ip(value: str | None) -> str:
    return (value or "").strip().replace(" ", "")


def normalize_endpoint_text(value: str | None) -> str:
    return re.sub(r"\.{2,}", ".", (value or "").strip().replace("：", ":"))


def normalize_capacity(value: str | None, unit: str) -> str:
    raw = (value or "").strip()
    if not raw:
        return ""
    raw = re.sub(rf"\s*{unit}\s*$", "", raw, flags=re.IGNORECASE)
    return raw.strip()


def split_endpoint_parts(value: str | None) -> list[str]:
    raw = normalize_endpoint_text(value)
    if not raw:
        return []
    parts = [part.strip() for part in re.split(r"[\r\n,，;；]+", raw) if part.strip()]
    ip_parts = [part for part in parts if extract_ip_like(part)]
    return ip_parts or parts


def endpoint_ip_and_ports(value: str | None) -> tuple[str, list[str]]:
    endpoint = normalize_endpoint_text(value)
    ip_value = extract_ip_like(endpoint)
    display = ip_value or endpoint
    endpoint_ports = extract_ports_freeform(endpoint) if ":" in endpoint else []
    return display, endpoint_ports


def short_mount_path(value: str | None) -> str:
    if not value:
        return ""
    raw = str(value).strip()
    if len(raw) <= 24:
        return raw
    parts = [p for p in raw.split("/") if p]
    if not parts:
        return raw
    if len(parts) == 1:
        return parts[0]
    return f".../{parts[-1]}"


def trim_text(value: str | None, limit: int) -> str:
    raw = (value or "").strip()
    if len(raw) <= limit:
        return raw
    if limit <= 1:
        return raw[:limit]
    return raw[: limit - 1] + "…"


def compact_ip(value: str | None) -> str:
    return trim_text(value or "", 18)


PNG_PIXEL = 9525
THIN_LINE = 9525
CARD_LINE = 12700
ZONE_LINE = 15240
LIGHT_LABEL_W = 400000
LIGHT_LABEL_H = 130000
WIRE_COLOR = "4A4A4A"
ZONE_BORDER = "2F2F2F"
TITLE_TEXT = "222222"
BODY_TEXT = "5A5A5A"
STANDARD_ICON_CM = 0.9
STANDARD_ICON_TEXT_SIZE = 600
SMALL_ICON_CM = 0.6
NODE_ICON_CM = 0.8

ICON_TARGETS = {
    "user": "../media/user-icon.png",
    "mobile": "../media/mobile-icon.png",
    "pc": "../media/pc-icon.png",
    "firewall": "../media/firewall-icon.png",
    "lb": "../media/lb-icon.png",
    "nginx": "../media/nginx-icon.png",
    "server": "../media/server-icon.png",
    "k8s": "../media/k8s-icon.png",
    "k8s_node": "../media/k8s-node-icon.png",
    "pod": "../media/pod-icon.png",
    "redis": "../media/redis-icon.png",
    "zookeeper": "../media/zookeeper-icon.png",
    "mq": "../media/mq-icon.png",
    "elk": "../media/elk-icon.png",
    "pg": "../media/db-icon.png",
    "mdd": "../media/mdd-icon.png",
    "nfs": "../media/nas-icon.png",
    "appstore": "../media/appstore-icon.png",
    "preview": "../media/preview-icon.png",
}

ICON_TARGET_TO_SOURCE = {
    ICON_TARGETS["user"]: ASSET_DIR / "用户.png",
    ICON_TARGETS["mobile"]: ASSET_DIR / "mobile.png",
    ICON_TARGETS["pc"]: ASSET_DIR / "pc.png",
    ICON_TARGETS["firewall"]: ASSET_DIR / "防火墙.png",
    ICON_TARGETS["lb"]: ASSET_DIR / "负载均衡.png",
    ICON_TARGETS["nginx"]: ASSET_DIR / "nginx.png",
    ICON_TARGETS["server"]: ASSET_DIR / "服务器.png",
    ICON_TARGETS["k8s"]: ASSET_DIR / "K8S.png",
    ICON_TARGETS["k8s_node"]: ASSET_DIR / "k8s-node.png",
    ICON_TARGETS["pod"]: ASSET_DIR / "k8S-pod.png",
    ICON_TARGETS["redis"]: ASSET_DIR / "redis.png",
    ICON_TARGETS["zookeeper"]: ASSET_DIR / "zookeeper.png",
    ICON_TARGETS["mq"]: ASSET_DIR / "mq 消息队列MQ.png",
    ICON_TARGETS["elk"]: ASSET_DIR / "elk.png",
    ICON_TARGETS["pg"]: ASSET_DIR / "数据库.png",
    ICON_TARGETS["mdd"]: ASSET_DIR / "多维数据库.png",
    ICON_TARGETS["nfs"]: ASSET_DIR / "NAS.png",
    ICON_TARGETS["appstore"]: ASSET_DIR / "app-store.png",
    ICON_TARGETS["preview"]: ASSET_DIR / "文件预览.png",
}

FAMILY_ICON_KEYS = {
    "lb": "lb",
    "nginx": "nginx",
    "gpaas": "k8s",
    "k8s": "k8s",
    "redis": "redis",
    "zookeeper": "zookeeper",
    "mq": "mq",
    "elk": "elk",
    "pg": "pg",
    "mdd": "mdd",
    "nfs": "nfs",
    "appstore": "appstore",
    "preview": "preview",
}


def resource_spec(resource: dict) -> str:
    parts = []
    if resource.get("cpu"):
        parts.append(f'{resource["cpu"]}C')
    if resource.get("memory"):
        parts.append(f'{resource["memory"]}G')
    if resource.get("data_disk"):
        parts.append(f'{resource["data_disk"]}G')
    return " ".join(parts)


def brief_resource_line(resource: dict) -> str:
    pieces = [resource.get("name") or resource.get("purpose") or "资源"]
    if resource.get("ip"):
        pieces.append(resource["ip"])
    spec = resource_spec(resource)
    if spec:
        pieces.append(spec)
    if resource.get("vip"):
        pieces.append(f'VIP {resource["vip"]}')
    if resource.get("lb_ip"):
        pieces.append(f'LB {resource["lb_ip"]}')
    return " | ".join([p for p in pieces if p])


def compact_port_labels(family: dict, limit: int = 3) -> list[str]:
    labels = []
    preferred_numeric = [trim_text(item, 10) for item in family.get("ports", []) if item]
    for item in preferred_numeric:
        labels.append(item)
        if len(labels) >= limit:
            return uniq_preserve(labels)
    for item in family.get("port_labels", []):
        labels.append(trim_text(item, 18))
        if len(labels) >= limit:
            break
    return uniq_preserve(labels)


def family_display_name(key: str, services: list[dict]) -> str:
    mapping = {
        "lb": "负载均衡",
        "nginx": "Nginx接入层",
        "gpaas": "gPaaS管理",
        "k8s": "K8S容器集群",
        "redis": "Redis集群",
        "zookeeper": "ZooKeeper",
        "mq": "MQ集群",
        "elk": "ELK日志服务",
        "pg": "关系数据库",
        "mdd": "多维数据库",
        "nfs": "共享存储(NFS)",
        "appstore": "AppStore",
        "preview": "文件预览",
    }
    if key in mapping:
        return mapping[key]
    if services and services[0].get("name"):
        return services[0]["name"]
    return key or "服务"


def split_lb_roles(family: dict | None) -> tuple[list[dict], dict | None]:
    if not family:
        return [], None
    external = []
    internal = None
    for resource in family.get("resources", []):
        merged = canonical_text(" ".join([resource.get("name", ""), resource.get("purpose", "")]))
        if "masterblb" in merged or ("内部" in (resource.get("purpose") or "") and "lb" in merged):
            internal = resource
        elif "移动" in (resource.get("name") or "") or "pc" in merged or "lb" in merged:
            external.append(resource)
    if not internal and external:
        for resource in list(external):
            merged = canonical_text(" ".join([resource.get("name", ""), resource.get("purpose", "")]))
            if "k8s" in merged:
                internal = resource
                external.remove(resource)
                break
    return external[:2], internal


def infer_resource_group(resource: dict) -> str:
    purpose = resource.get("purpose", "")
    name = resource.get("name", "")
    merged = canonical_text(" ".join([purpose, name, resource.get("resource_type", "")]))
    if "lb" in merged or "blb" in merged or "负载均衡" in purpose:
        return "lb"
    if "gpaas" in merged:
        return "gpaas"
    if "容器集群" in purpose or "k8s" in merged or "cce" in merged:
        return "k8s"
    if "接入服务" in purpose or "ng" in canonical_text(name):
        return "nginx"
    if "多维" in purpose or "mdd" in merged:
        return "mdd"
    if "zookeeper" in merged or "zk" in merged or canonical_text(name).startswith("ierppzk"):
        return "zookeeper"
    if "redis" in merged:
        return "redis"
    if "mq" in merged or "admq" in merged or "rabbitmq" in merged:
        return "mq"
    if "elk" in merged or "elasticsearch" in merged or "kafka" in merged or "logstash" in merged or "es" in canonical_text(purpose):
        return "elk"
    if "达梦数据库" in purpose or "postgresql" in merged or re.search(r"\bpg\d|\bdm\d|\bdb\d", canonical_text(name)):
        return "pg"
    if "文件预览" in purpose or "preview" in merged:
        return "preview"
    if "nfs" in merged or "共享存储" in merged:
        return "nfs"
    if "rpa" in merged or "robot" in merged or "机器人" in merged:
        return "other"
    return ""


def service_family_key(service: dict) -> str:
    merged = canonical_text(" ".join([service.get("name", ""), service.get("purpose", ""), service.get("category", "")]))
    if merged.startswith("lb") or "负载均衡" in service.get("purpose", ""):
        return "lb"
    if "nginx" in merged:
        return "nginx"
    if "gpaas" in merged or "容器管理平台" in service.get("name", ""):
        return "gpaas"
    if "k8s" in merged or "容器服务" in service.get("name", ""):
        return "k8s"
    if "redis" in merged:
        return "redis"
    if "zookeer" in merged or "zookeeper" in merged:
        return "zookeeper"
    if "admq" in merged or "消息队列" in service.get("purpose", ""):
        return "mq"
    if any(token in merged for token in ("kafka", "logstash", "elasticsearch", "日志组件")):
        return "elk"
    if merged.startswith("pg") or "关系数据库" in service.get("purpose", ""):
        return "pg"
    if merged.startswith("mdd") or "多维" in service.get("name", ""):
        return "mdd"
    if "共享存储" in service.get("name", "") or "nfs" in merged:
        return "nfs"
    if "appstore" in merged:
        return "appstore"
    return merged or "other"


def zone_for_family(key: str, category: str = "") -> str:
    if key in {"lb", "nginx"}:
        return "access"
    if key in {"gpaas", "k8s", "preview"}:
        return "application"
    if key in {"pg", "mdd"} or "数据库" in category:
        return "data"
    return "platform"


def letters_to_index(letters: str) -> int:
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch.upper()) - 64)
    return value


class XlsxReader:
    def __init__(self, path: str | Path):
        self.path = Path(path).resolve()
        self._zip = zipfile.ZipFile(self.path)
        self.shared_strings = self._load_shared_strings()
        self.sheets = self._load_sheet_targets()

    def close(self) -> None:
        self._zip.close()

    def _load_shared_strings(self) -> list[str]:
        if "xl/sharedStrings.xml" not in self._zip.namelist():
            return []
        root = ET.fromstring(self._zip.read("xl/sharedStrings.xml"))
        values = []
        for item in root.findall("a:si", NS_XLS):
            text = "".join(node.text or "" for node in item.iterfind(".//a:t", NS_XLS))
            values.append(text)
        return values

    def _load_sheet_targets(self) -> dict[str, str]:
        workbook = ET.fromstring(self._zip.read("xl/workbook.xml"))
        rels = ET.fromstring(self._zip.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall("p:Relationship", NS_XLS)}
        sheets = {}
        for sheet in workbook.find("a:sheets", NS_XLS):
            rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            sheets[sheet.attrib["name"]] = "xl/" + rel_map[rid]
        return sheets

    def read_sheet_matrix(self, sheet_name: str) -> list[list[str]]:
        if sheet_name not in self.sheets:
            raise KeyError(f"Missing sheet: {sheet_name}")
        root = ET.fromstring(self._zip.read(self.sheets[sheet_name]))
        rows = []
        max_width = 0
        for row in root.findall(".//a:sheetData/a:row", NS_XLS):
            cells = {}
            for cell in row.findall("a:c", NS_XLS):
                ref = cell.attrib.get("r", "")
                letters = "".join(ch for ch in ref if ch.isalpha())
                if not letters:
                    continue
                col_idx = letters_to_index(letters)
                value = ""
                cell_type = cell.attrib.get("t")
                if cell_type == "s":
                    node = cell.find("a:v", NS_XLS)
                    if node is not None and node.text is not None:
                        value = self.shared_strings[int(node.text)]
                elif cell_type == "inlineStr":
                    node = cell.find("a:is", NS_XLS)
                    if node is not None:
                        value = "".join(text.text or "" for text in node.iterfind(".//a:t", NS_XLS))
                else:
                    node = cell.find("a:v", NS_XLS)
                    if node is not None and node.text is not None:
                        value = node.text
                cells[col_idx] = value
                max_width = max(max_width, col_idx)
            rows.append(cells)

        matrix = []
        for row in rows:
            values = [""] * max_width
            for idx, value in row.items():
                values[idx - 1] = value
            matrix.append(values)
        return matrix


def get_cell(row: list[str], index: int) -> str:
    if index < len(row):
        return str(row[index]).strip()
    return ""


def parse_resource_sheet(rows: list[list[str]]) -> list[dict]:
    if matrix_contains(rows, ["资产名称", "IP", "CPU", "内存"]):
        return parse_simple_resource_sheet(rows)

    resources = []
    current_env = ""
    current_purpose = ""
    for row in rows[3:]:
        env = get_cell(row, 0)
        purpose = get_cell(row, 1)
        if env:
            current_env = env
        if purpose:
            current_purpose = purpose

        raw_fields = {
            "env": env,
            "purpose": purpose,
            "resource_type": get_cell(row, 2),
            "name": get_cell(row, 3),
            "os_version": get_cell(row, 4),
            "ip": get_cell(row, 5),
            "vip": get_cell(row, 7),
            "lb_ip": get_cell(row, 8),
            "cpu": get_cell(row, 9),
            "memory": get_cell(row, 10),
            "system_disk": get_cell(row, 11),
            "data_disk": get_cell(row, 12),
            "mount_dir": get_cell(row, 13),
            "disk_type": get_cell(row, 14),
        }
        meaningful_direct_fields = (
            raw_fields["name"],
            raw_fields["ip"],
            raw_fields["vip"],
            raw_fields["lb_ip"],
            raw_fields["cpu"],
            raw_fields["memory"],
            raw_fields["system_disk"],
            raw_fields["data_disk"],
            raw_fields["mount_dir"],
        )
        if not any(meaningful_direct_fields):
            continue

        record = {
            "env": current_env,
            "purpose": current_purpose,
            "resource_type": raw_fields["resource_type"],
            "name": raw_fields["name"],
            "os_version": raw_fields["os_version"],
            "ip": raw_fields["ip"],
            "vip": raw_fields["vip"],
            "lb_ip": raw_fields["lb_ip"],
            "cpu": raw_fields["cpu"],
            "memory": raw_fields["memory"],
            "system_disk": raw_fields["system_disk"],
            "data_disk": raw_fields["data_disk"],
            "mount_dir": raw_fields["mount_dir"],
            "disk_type": raw_fields["disk_type"],
        }

        record["group_hint"] = infer_resource_group(record)
        record["last_octet"] = extract_last_octet(record["ip"])
        resources.append(record)
    return resources


def parse_simple_resource_sheet(rows: list[list[str]]) -> list[dict]:
    resources = []
    current_env = ""
    current_name = ""
    header = rows[0] if rows else []
    header_map = {canonical_text(value): index for index, value in enumerate(header)}

    env_idx = header_map.get(canonical_text("环境"), 0)
    name_idx = header_map.get(canonical_text("资产名称"), 1)
    type_idx = header_map.get(canonical_text("资源类型"), 2)
    ip_idx = header_map.get(canonical_text("IP"), 3)
    port_idx = header_map.get(canonical_text("端口"), 4)
    cpu_idx = header_map.get(canonical_text("CPU"), 5)
    memory_idx = header_map.get(canonical_text("内存"), 6)
    system_disk_idx = header_map.get(canonical_text("系统盘"), 7)
    data_disk_idx = header_map.get(canonical_text("数据盘"), 8)
    remark_idx = header_map.get(canonical_text("备注"), 10)

    for row in rows[1:]:
        env = get_cell(row, env_idx)
        asset_name = get_cell(row, name_idx)
        resource_type = get_cell(row, type_idx)
        endpoints = split_endpoint_parts(get_cell(row, ip_idx))
        ports = parse_port_entries(get_cell(row, port_idx))
        cpu = normalize_capacity(get_cell(row, cpu_idx), "C")
        memory = normalize_capacity(get_cell(row, memory_idx), "G")
        system_disk = normalize_capacity(get_cell(row, system_disk_idx), "G")
        data_disk = normalize_capacity(get_cell(row, data_disk_idx), "G")
        remark = get_cell(row, remark_idx)

        if env:
            current_env = env
        if asset_name and not endpoints:
            current_name = asset_name
            continue
        if asset_name:
            current_name = asset_name
        if not endpoints:
            continue

        for index, endpoint in enumerate(endpoints):
            ip_value, endpoint_ports = endpoint_ip_and_ports(endpoint)
            all_ports = uniq_preserve([entry["port"] for entry in ports if entry.get("port")] + endpoint_ports)
            base_name = current_name or asset_name or "服务器"
            name = base_name
            if len(endpoints) > 1:
                name = f"{base_name}-{index + 1}"
            record = {
                "env": current_env,
                "purpose": " ".join([base_name, remark]).strip(),
                "resource_type": resource_type or "服务器",
                "name": name,
                "os_version": "",
                "ip": ip_value,
                "vip": "",
                "lb_ip": "",
                "cpu": cpu,
                "memory": memory,
                "system_disk": system_disk,
                "data_disk": data_disk,
                "mount_dir": "",
                "disk_type": "",
                "ports": all_ports,
                "group_hint": "",
                "last_octet": extract_last_octet(ip_value),
            }
            record["group_hint"] = infer_resource_group(record)
            resources.append(record)
    return resources


def parse_service_resource_refs(raw: str) -> list[dict]:
    refs = []
    for line in split_nonempty_lines(raw):
        label = ""
        endpoint = line
        if ":" in line or "：" in line:
            left, right = re.split(r"[:：]", line, 1)
            if extract_ip_like(left) is None and not left.lower().startswith("xx.xx"):
                label = left.strip()
                endpoint = right.strip()
        ip_value = extract_ip_like(endpoint) or extract_ip_like(line)
        refs.append(
            {
                "raw": line,
                "label": label,
                "endpoint": endpoint.strip(),
                "ip": ip_value,
                "last_octet": extract_last_octet(ip_value or endpoint),
            }
        )
    return refs


def parse_port_entries(raw: str) -> list[dict]:
    entries = []
    for line in split_nonempty_lines(raw):
        normalized = line.replace("：", ":")
        match = re.match(r"\s*([^:]+)\s*:\s*(\d{2,5})\s*$", normalized)
        if match:
            entries.append({"label": match.group(1).strip(), "port": match.group(2), "raw": line})
            continue
        numbers = re.findall(r"\b(\d{2,5})\b", normalized)
        if numbers:
            for number in numbers:
                entries.append({"label": "", "port": number, "raw": line})
        else:
            entries.append({"label": "", "port": "", "raw": line})
    return uniq_preserve(entries)


def parse_service_sheet(rows: list[list[str]]) -> list[dict]:
    services = []
    current_category = ""
    for row in rows[2:]:
        category = get_cell(row, 0)
        if category:
            current_category = category
        name = get_cell(row, 1)
        if not name:
            continue
        remark = get_cell(row, 8)
        service = {
            "category": current_category,
            "name": name,
            "purpose": get_cell(row, 2),
            "version": get_cell(row, 3),
            "deploy_mode": get_cell(row, 4),
            "source": get_cell(row, 5),
            "service_resources_raw": get_cell(row, 6),
            "access_ports_raw": get_cell(row, 7),
            "remark": remark,
            "resource_refs": parse_service_resource_refs(get_cell(row, 6)),
            "access_ports": parse_port_entries(get_cell(row, 7)),
            "remark_ports": extract_ports_freeform(remark),
        }
        vip_match = re.search(r"VIP[:：]\s*([^\s（(]+)", remark)
        service["vip_hint"] = vip_match.group(1) if vip_match else ""
        service["family_key"] = service_family_key(service)
        services.append(service)
    return services


def parse_pod_sheet(rows: list[list[str]]) -> list[dict]:
    pods = []
    current_env = ""
    for row in rows[2:]:
        env = get_cell(row, 0)
        if env:
            current_env = env
        name = get_cell(row, 1)
        if not name:
            continue
        if "资源合计" in name:
            break
        pods.append(
            {
                "env": current_env,
                "name": name,
                "description": get_cell(row, 2),
                "cpu": get_cell(row, 3),
                "memory": get_cell(row, 4),
                "replicas": get_cell(row, 5),
                "total_cpu": get_cell(row, 6),
                "total_memory": get_cell(row, 7),
                "jvm_heap": get_cell(row, 8),
                "container_port": get_cell(row, 9),
                "external_port": get_cell(row, 10),
                "container_path": get_cell(row, 11),
                "host_path": get_cell(row, 12),
                "domain": get_cell(row, 13),
                "appids": get_cell(row, 14),
                "libs": get_cell(row, 15),
            }
        )
    return pods


def build_resource_indexes(resources: list[dict]) -> dict:
    by_ip = defaultdict(list)
    by_last_octet = defaultdict(list)
    by_name = defaultdict(list)
    by_group = defaultdict(list)

    for resource in resources:
        ip_value = normalize_ip(resource.get("ip"))
        if ip_value:
            by_ip[ip_value].append(resource)
        last_octet = resource.get("last_octet")
        if last_octet:
            by_last_octet[last_octet].append(resource)
        for key in filter(None, [canonical_text(resource.get("name")), canonical_text(resource.get("purpose"))]):
            by_name[key].append(resource)
        if resource.get("group_hint"):
            by_group[resource["group_hint"]].append(resource)

    return {
        "by_ip": by_ip,
        "by_last_octet": by_last_octet,
        "by_name": by_name,
        "by_group": by_group,
    }


def dedupe_resources(resources: Iterable[dict]) -> list[dict]:
    seen = set()
    result = []
    for resource in resources:
        key = (resource.get("name"), resource.get("ip"), resource.get("purpose"))
        if key in seen:
            continue
        seen.add(key)
        result.append(resource)
    return result


def match_resources_to_service(service: dict, indexes: dict) -> list[dict]:
    if service.get("direct_resources") is not None:
        return dedupe_resources(service["direct_resources"])

    matched = []
    for ref in service.get("resource_refs", []):
        ref_matches = []
        if ref.get("ip"):
            ref_matches.extend(indexes["by_ip"].get(normalize_ip(ref["ip"]), []))
        if not ref_matches and ref.get("last_octet"):
            ref_matches.extend(indexes["by_last_octet"].get(ref["last_octet"], []))
        label_key = canonical_text(ref.get("label") or "")
        if label_key:
            for key, values in indexes["by_name"].items():
                if label_key in key or key in label_key:
                    ref_matches.extend(values)
        matched.extend(ref_matches)
    if not matched and service.get("family_key"):
        matched.extend(indexes["by_group"].get(service["family_key"], []))
    return dedupe_resources(matched)


def build_families(resources: list[dict], services: list[dict], pods: list[dict]) -> tuple[dict[str, dict], list[dict]]:
    indexes = build_resource_indexes(resources)
    families: dict[str, dict] = {}
    all_matched_resource_keys = set()

    grouped_services = defaultdict(list)
    for service in services:
        grouped_services[service["family_key"]].append(service)

    for key, family_services in grouped_services.items():
        family_resources = []
        for service in family_services:
            matched = match_resources_to_service(service, indexes)
            service["matched_resources"] = matched
            for resource in matched:
                all_matched_resource_keys.add((resource.get("name"), resource.get("ip"), resource.get("purpose")))
            family_resources.extend(matched)

        display_ports = []
        numeric_ports = []
        raw_endpoints = []
        notes = []
        for service in family_services:
            for entry in service.get("access_ports", []):
                if entry.get("port"):
                    numeric_ports.append(entry["port"])
                    display_ports.append(f'{entry["label"]}:{entry["port"]}' if entry.get("label") else entry["port"])
            numeric_ports.extend(service.get("remark_ports", []))
            numeric_ports.extend(extract_ports_freeform(service.get("service_resources_raw")))
            notes.extend(split_nonempty_lines(service.get("remark")))
            raw_endpoints.extend([ref["raw"] for ref in service.get("resource_refs", []) if ref.get("raw")])

        for resource in family_resources:
            numeric_ports.extend(resource.get("ports", []))

        if key == "k8s":
            for pod in pods:
                if pod.get("external_port"):
                    numeric_ports.append(pod["external_port"])
                    display_ports.append(f'{pod["name"]}:{pod["external_port"]}')

        family = {
            "key": key,
            "display_name": family_display_name(key, family_services),
            "zone": zone_for_family(key, family_services[0].get("category", "")),
            "services": family_services,
            "resources": dedupe_resources(family_resources),
            "ports": uniq_preserve([port for port in numeric_ports if port]),
            "port_labels": uniq_preserve([port for port in display_ports if port]),
            "notes": uniq_preserve([note for note in notes if note]),
            "raw_endpoints": uniq_preserve([item for item in raw_endpoints if item]),
        }
        families[key] = family

    for extra_key in ("preview",):
        if extra_key in families:
            continue
        group_resources = indexes["by_group"].get(extra_key, [])
        if not group_resources:
            continue
        for resource in group_resources:
            all_matched_resource_keys.add((resource.get("name"), resource.get("ip"), resource.get("purpose")))
        families[extra_key] = {
            "key": extra_key,
            "display_name": family_display_name(extra_key, []),
            "zone": zone_for_family(extra_key),
            "services": [],
            "resources": dedupe_resources(group_resources),
            "ports": [],
            "port_labels": [],
            "notes": [],
            "raw_endpoints": [],
        }

    unmatched = []
    for resource in resources:
        key = (resource.get("name"), resource.get("ip"), resource.get("purpose"))
        if key not in all_matched_resource_keys and resource.get("group_hint") not in {"gpaas", "k8s", "nginx", "mdd", "pg", "redis", "zookeeper", "mq", "elk", "preview"}:
            unmatched.append(resource)

    return families, unmatched


def detect_port_conflicts(services: list[dict], pods: list[dict]) -> list[dict]:
    conflicts = []
    service_ports = {}
    for service in services:
        if service.get("family_key") != "k8s":
            continue
        for entry in service.get("access_ports", []):
            if entry.get("label") and entry.get("port"):
                service_ports[canonical_text(entry["label"])] = entry["port"]

    for pod in pods:
        pod_key = canonical_text(pod.get("name"))
        pod_port = pod.get("external_port")
        if pod_key in service_ports and pod_port and service_ports[pod_key] != pod_port:
            conflicts.append(
                {
                    "service": pod.get("name"),
                    "service_sheet_port": service_ports[pod_key],
                    "pod_sheet_port": pod_port,
                }
            )
    return conflicts


def matrix_contains(rows: list[list[str]], keywords: list[str], sample_rows: int = 6) -> bool:
    haystack = canonical_text(" ".join(" ".join(row) for row in rows[:sample_rows]))
    return all(canonical_text(keyword) in haystack for keyword in keywords)


def identify_sheet_roles(reader: XlsxReader) -> dict[str, str]:
    roles: dict[str, str] = {}
    for sheet_name in reader.sheets:
        rows = reader.read_sheet_matrix(sheet_name)
        if not rows:
            continue
        if "resource" not in roles and (
            matrix_contains(rows, ["机器名/服务名", "IP/域名", "数据盘"])
            or matrix_contains(rows, ["资产名称", "IP", "CPU", "内存"])
        ):
            roles["resource"] = sheet_name
            continue
        if "service" not in roles and matrix_contains(rows, ["服务名称", "服务资源", "服务访问端口"]):
            roles["service"] = sheet_name
            continue
        if "pod" not in roles and matrix_contains(rows, ["容器节点名称", "容器端口", "外部端口"]):
            roles["pod"] = sheet_name
            continue
    return roles


def synthesize_services_from_resources(resources: list[dict], pods: list[dict]) -> list[dict]:
    grouped: dict[str, list[dict]] = defaultdict(list)
    lb_ips = []
    for resource in resources:
        if resource.get("group_hint"):
            grouped[resource["group_hint"]].append(resource)
        if resource.get("lb_ip"):
            lb_ips.append(resource["lb_ip"])

    services: list[dict] = []

    def add_service(
        family_key: str,
        name: str,
        purpose: str,
        category: str,
        deploy_mode: str = "",
        access_ports: list[dict] | None = None,
        remark: str = "",
        vip_hint: str = "",
    ) -> None:
        resources_for_family = dedupe_resources(grouped.get(family_key, []))
        if not resources_for_family and family_key not in {"lb", "nfs", "appstore"}:
            return
        service = {
            "category": category,
            "name": name,
            "purpose": purpose,
            "version": "",
            "deploy_mode": deploy_mode,
            "source": "",
            "direct_resources": resources_for_family,
            "service_resources_raw": "\n".join([resource.get("ip") or resource.get("name") or "" for resource in resources_for_family if resource.get("ip") or resource.get("name")]),
            "access_ports_raw": "\n".join(
                [f'{entry["label"]}:{entry["port"]}' if entry.get("label") else entry["port"] for entry in (access_ports or []) if entry.get("port")]
            ),
            "remark": remark,
            "resource_refs": [{"raw": resource.get("ip") or resource.get("name") or "", "label": "", "endpoint": resource.get("ip") or resource.get("name") or "", "ip": resource.get("ip"), "last_octet": resource.get("last_octet")} for resource in resources_for_family],
            "access_ports": uniq_preserve(access_ports or []),
            "remark_ports": extract_ports_freeform(remark),
            "vip_hint": vip_hint,
            "family_key": family_key,
        }
        services.append(service)

    if lb_ips or grouped.get("lb"):
        add_service(
            "lb",
            "LB1",
            "负载均衡",
            "接入",
            access_ports=[{"label": "", "port": "80", "raw": "80"}],
            remark=f'LB-IP {lb_ips[0]}' if lb_ips else "",
        )

    add_service("nginx", "Nginx", "反向代理/静态资源服务/应用仓库", "接入", deploy_mode="集群")
    add_service("gpaas", "容器管理平台", "gPaas", "容器", access_ports=[{"label": "", "port": "8060", "raw": "8060"}])

    k8s_ports = []
    for pod in pods:
        if pod.get("external_port"):
            k8s_ports.append({"label": pod.get("name", ""), "port": pod["external_port"], "raw": f'{pod.get("name","")}:{pod["external_port"]}'})
    api_vip = next((resource.get("vip") for resource in grouped.get("k8s", []) if resource.get("vip")), "")
    add_service(
        "k8s",
        "k8s容器服务",
        "",
        "容器",
        access_ports=uniq_preserve(k8s_ports),
        remark=f'6443端口通过VIP：{api_vip}（master节点vip地址）' if api_vip else "",
        vip_hint=api_vip,
    )

    add_service("redis", "Redis", "分布式缓存", "平台")
    add_service("zookeeper", "ZooKeeper", "注册与发现", "平台")
    add_service("mq", "admq", "消息队列", "平台")
    add_service("elk", "ELK", "日志组件", "平台")
    add_service("pg", "pg", "关系数据库", "数据")
    add_service("mdd", "mdd", "", "数据")
    add_service("preview", "文件预览", "", "应用")

    if grouped.get("preview") or grouped.get("nginx"):
        add_service("nfs", "共享存储（NFS）", "", "平台", remark="容器持久化（图片附件，轻分析数据）；appstore、静态资源")
        add_service("appstore", "Appstore", "", "平台")

    return services


def read_template_slide_size(template_path: Path) -> tuple[int, int]:
    with zipfile.ZipFile(template_path) as archive:
        root = ET.fromstring(archive.read("ppt/presentation.xml"))
    ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
    node = root.find("p:sldSz", ns)
    if node is None:
        return BASE_SLIDE_W, BASE_SLIDE_H
    return int(node.attrib["cx"]), int(node.attrib["cy"])


class SlideBuilder:
    def __init__(self, width: int, height: int):
        self.width = width
        self.height = height
        self._next_id = 2
        self.parts: list[str] = []
        self.text_margin = 36000
        self.image_rels: list[tuple[str, str]] = []

    def sx(self, value: int) -> int:
        return int(value * self.width / BASE_SLIDE_W)

    def sy(self, value: int) -> int:
        return int(value * self.height / BASE_SLIDE_H)

    def cm(self, value: float) -> int:
        return int(round(value * 360000))

    def _new_id(self) -> int:
        current = self._next_id
        self._next_id += 1
        return current

    def _solid_fill(self, color: str | None) -> str:
        if not color:
            return "<a:noFill/>"
        return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'

    def _line_xml(self, color: str | None, width: int, dash: str | None = None) -> str:
        if not color:
            return "<a:ln><a:noFill/></a:ln>"
        dash_xml = f'<a:prstDash val="{dash}"/>' if dash else ""
        return f'<a:ln w="{width}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>{dash_xml}</a:ln>'

    def _paragraph_xml(self, text: str, size: int, color: str, bold: bool = False, align: str = "ctr") -> str:
        bold_flag = ' b="1"' if bold else ""
        safe_text = escape(text)
        return (
            f'<a:p>'
            f'<a:pPr algn="{align}"><a:defRPr sz="{size}"{bold_flag}><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:defRPr></a:pPr>'
            f'<a:r><a:rPr lang="zh-CN" altLang="en-US" sz="{size}"{bold_flag}><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="微软雅黑"/><a:ea typeface="微软雅黑"/></a:rPr><a:t>{safe_text}</a:t></a:r>'
            f'<a:endParaRPr lang="zh-CN" altLang="en-US" sz="{size}"{bold_flag}/>'
            f'</a:p>'
        )

    def _text_body(self, paragraphs: list[dict] | None, anchor: str = "ctr", autofit: str = "noAutofit") -> str:
        if not paragraphs:
            paragraphs = [{"text": "", "size": 1000, "color": "000000", "bold": False, "align": "ctr"}]
        xml = [
            f'<p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="{anchor}" '
            f'lIns="{self.text_margin}" rIns="{self.text_margin}" tIns="{self.text_margin}" bIns="{self.text_margin}">'
            f'<a:{autofit}/></a:bodyPr><a:lstStyle/>'
        ]
        for paragraph in paragraphs:
            xml.append(
                self._paragraph_xml(
                    paragraph.get("text", ""),
                    int(paragraph.get("size", 1000)),
                    paragraph.get("color", "000000"),
                    bool(paragraph.get("bold", False)),
                    paragraph.get("align", "ctr"),
                )
            )
        xml.append("</p:txBody>")
        return "".join(xml)

    def add_round_rect(
        self,
        x: int,
        y: int,
        cx: int,
        cy: int,
        fill: str | None,
        border: str | None,
        paragraphs: list[dict] | None = None,
        border_width: int = CARD_LINE,
        name: str = "Rounded Rectangle",
        shape: str = "roundRect",
        dash: str | None = None,
    ) -> None:
        shape_id = self._new_id()
        self.parts.append(
            f'<p:sp>'
            f'<p:nvSpPr><p:cNvPr id="{shape_id}" name="{escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="{shape}"><a:avLst/></a:prstGeom>{self._solid_fill(fill)}{self._line_xml(border, border_width, dash=dash)}</p:spPr>'
            f'{self._text_body(paragraphs)}</p:sp>'
        )

    def add_text_box(
        self,
        x: int,
        y: int,
        cx: int,
        cy: int,
        paragraphs: list[dict],
        name: str = "TextBox",
        wrap: str = "square",
    ) -> None:
        shape_id = self._new_id()
        text_body = self._text_body(paragraphs).replace('wrap="square"', f'wrap="{wrap}"', 1)
        self.parts.append(
            f'<p:sp>'
            f'<p:nvSpPr><p:cNvPr id="{shape_id}" name="{escape(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'{text_body}</p:sp>'
        )

    def add_line_label(
        self,
        x1: int,
        y1: int,
        x2: int,
        y2: int,
        text: str,
        color: str,
        w: int | None = None,
        h: int | None = None,
        x_offset: int = 0,
        y_offset: int = 0,
    ) -> None:
        label_w = w or self.sx(LIGHT_LABEL_W)
        label_h = h or self.sy(LIGHT_LABEL_H)
        mid_x = (x1 + x2) // 2 - label_w // 2 + x_offset
        mid_y = (y1 + y2) // 2 - label_h // 2 - self.sy(90000) + y_offset
        self.add_round_rect(
            mid_x,
            mid_y,
            label_w,
            label_h,
            "FFFFFF",
            None,
            [{"text": trim_text(text, 14), "size": 460, "color": WIRE_COLOR, "bold": False}],
            border_width=THIN_LINE,
            name="LineLabel",
        )

    def add_connector(
        self,
        x1: int,
        y1: int,
        x2: int,
        y2: int,
        color: str,
        width: int = THIN_LINE,
        arrow: bool = True,
        name: str | None = None,
    ) -> None:
        shape_id = self._new_id()
        off_x = min(x1, x2)
        off_y = min(y1, y2)
        ext_x = max(abs(x2 - x1), 12700)
        ext_y = max(abs(y2 - y1), 12700)
        flip_h = ' flipH="1"' if x1 > x2 else ""
        flip_v = ' flipV="1"' if y1 > y2 else ""
        tail_end = '<a:tailEnd type="stealth"/>' if arrow else ""
        shape_name = escape(name or f"Connector {shape_id}")
        self.parts.append(
            f'<p:cxnSp>'
            f'<p:nvCxnSpPr><p:cNvPr id="{shape_id}" name="{shape_name}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
            f'<p:spPr><a:xfrm{flip_h}{flip_v}><a:off x="{off_x}" y="{off_y}"/><a:ext cx="{ext_x}" cy="{ext_y}"/></a:xfrm>'
            f'<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
            f'<a:ln w="{width}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>{tail_end}</a:ln></p:spPr>'
            f'</p:cxnSp>'
        )

    def add_vertical_line(self, x: int, y1: int, y2: int, color: str, width: int = THIN_LINE, arrow: bool = True, name: str = "VerticalLine") -> None:
        top = min(y1, y2)
        height = max(abs(y2 - y1), 12700)
        tail_end = '<a:tailEnd type="stealth"/>' if arrow else ""
        shape_id = self._new_id()
        self.parts.append(
            f'<p:sp>'
            f'<p:nvSpPr><p:cNvPr id="{shape_id}" name="{escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>'
            f'<a:ln w="{width}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>{tail_end}</a:ln></p:spPr>'
            f'</p:sp>'
        )

    def add_image(self, x: int, y: int, cx: int, cy: int, target: str, name: str = "Picture") -> None:
        shape_id = self._new_id()
        rel_id = f"rIdImg{len(self.image_rels) + 1}"
        self.image_rels.append((rel_id, target))
        self.parts.append(
            f'<p:pic>'
            f'<p:nvPicPr><p:cNvPr id="{shape_id}" name="{escape(name)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rel_id}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
            f'</p:pic>'
        )

    def add_icon_card(
        self,
        x: int,
        y: int,
        cx: int,
        cy: int,
        icon_target: str,
        detail_lines: list[str],
        border: str,
        fill: str = "FFFFFF",
        name: str = "IconCard",
    ) -> None:
        self.add_round_rect(x, y, cx, cy, fill, border, paragraphs=None, border_width=THIN_LINE, name=f"{name}Shell")
        icon_box = min(cx - self.sx(100000), cy - self.sy(180000), self.sy(330000))
        icon_x = x + (cx - icon_box) // 2
        icon_y = y + self.sy(40000)
        self.add_image(icon_x, icon_y, icon_box, icon_box, icon_target, name=f"{name}Icon")
        text_y = icon_y + icon_box + self.sy(20000)
        text_h = max(self.sy(120000), y + cy - text_y - self.sy(30000))
        paragraphs = [{"text": trim_text(line, 18), "size": 500, "color": "495057", "bold": False} for line in detail_lines if line]
        if not paragraphs:
            paragraphs = [{"text": "-", "size": 520, "color": "495057"}]
        self.add_text_box(x + self.sx(30000), text_y, cx - self.sx(60000), text_h, paragraphs, name=f"{name}Text")

    def add_icon_label(
        self,
        x: int,
        y: int,
        cx: int,
        cy: int,
        icon_target: str,
        detail_lines: list[str],
        name: str = "IconLabel",
        text_color: str = BODY_TEXT,
        icon_cap: int | None = None,
        text_size: int = 500,
        frame: bool = False,
        frame_color: str = ZONE_BORDER,
    ) -> None:
        if frame:
            self.add_round_rect(
                x,
                y,
                cx,
                cy,
                None,
                frame_color,
                paragraphs=None,
                border_width=THIN_LINE,
                name=f"{name}Frame",
                shape="rect",
                dash="dash",
            )
        is_pod_icon = icon_target == ICON_TARGETS["pod"]
        if icon_target == ICON_TARGETS["server"]:
            standard_icon = self.cm(STANDARD_ICON_CM)
        elif icon_target in {ICON_TARGETS["mobile"], ICON_TARGETS["pc"], ICON_TARGETS["k8s"]}:
            standard_icon = self.cm(SMALL_ICON_CM)
        elif icon_target == ICON_TARGETS["k8s_node"]:
            standard_icon = self.cm(NODE_ICON_CM)
        elif is_pod_icon:
            standard_icon = self.sy(210000)
        else:
            standard_icon = self.cm(STANDARD_ICON_CM)
        icon_cap_value = (icon_cap or standard_icon) if is_pod_icon else standard_icon
        icon_box = min(cx, max(standard_icon, cy - self.sy(180000)), icon_cap_value)
        icon_x = x + (cx - icon_box) // 2
        icon_y = y + self.sy(10000)
        self.add_image(icon_x, icon_y, icon_box, icon_box, icon_target, name=f"{name}Icon")
        text_y = icon_y + icon_box + self.sy(10000)
        text_h = max(self.sy(100000), y + cy - text_y)
        effective_text_size = max(text_size, STANDARD_ICON_TEXT_SIZE)
        paragraphs = [{"text": trim_text(line, 18), "size": effective_text_size, "color": text_color, "bold": False} for line in detail_lines if line]
        if not paragraphs:
            paragraphs = [{"text": "-", "size": effective_text_size, "color": text_color}]
        self.add_text_box(x, text_y, cx, text_h, paragraphs, name=f"{name}Text", wrap="none")

    def build(self) -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            + "".join(self.parts)
            + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
        )


def family_paragraphs(family: dict, title_color: str, body_color: str) -> list[dict]:
    lines = [{"text": family["display_name"], "size": 1050, "color": title_color, "bold": True}]

    if family.get("port_labels"):
        sample = "、".join(family["port_labels"][:3])
        lines.append({"text": f"端口 {sample}", "size": 760, "color": body_color})
    elif family.get("ports"):
        lines.append({"text": f'端口 {",".join(family["ports"][:4])}', "size": 760, "color": body_color})

    resource_lines = [brief_resource_line(resource) for resource in family.get("resources", [])][:3]
    if resource_lines:
        for line in resource_lines:
            lines.append({"text": line, "size": 680, "color": body_color})
    elif family.get("raw_endpoints"):
        lines.append({"text": family["raw_endpoints"][0], "size": 680, "color": body_color})

    if family.get("notes"):
        lines.append({"text": family["notes"][0], "size": 650, "color": body_color})

    return lines[:6]


def server_box_paragraphs(
    resource: dict,
    port_labels: list[str],
    title_color: str,
    body_color: str,
    small: bool = False,
    include_ports: bool = True,
) -> list[dict]:
    title_size = 760 if small else 820
    body_size = 620 if small else 680
    spec_size = 560 if small else 620
    lines = [
        {"text": trim_text(resource.get("name") or resource.get("purpose") or "服务器", 18), "size": title_size, "color": title_color, "bold": True},
        {"text": compact_ip(resource.get("ip") or "-"), "size": body_size, "color": body_color},
    ]
    if include_ports and port_labels:
        lines.append({"text": f'端口 {trim_text(",".join(port_labels[:2]), 18)}', "size": body_size, "color": body_color})
    spec = resource_spec(resource)
    if spec:
        lines.append({"text": trim_text(spec, 24 if small else 28), "size": spec_size, "color": body_color})
    return lines[:4]


def icon_detail_lines(resource: dict, only_ip: bool = False) -> list[str]:
    lines = []
    ip_text = compact_ip(resource.get("ip") or "-")
    if ip_text:
        lines.append(ip_text)
    if not only_ip:
        spec = resource_spec(resource)
        if spec:
            lines.append(trim_text(spec, 18))
    return lines[:2]


def resource_icon_target(resource: dict, fallback: str) -> str:
    kind = canonical_text(resource.get("resource_type", ""))
    if "云服务" in (resource.get("resource_type") or ""):
        return ICON_TARGETS.get(FAMILY_ICON_KEYS.get(fallback, fallback), ICON_TARGETS["server"])
    if any(token in kind for token in ("服务器", "虚拟机", "ecs主机", "ecs", "主机")):
        return ICON_TARGETS["server"]
    return ICON_TARGETS["server"]


def pod_icon_lines(pod: dict) -> list[str]:
    lines = []
    title = pod.get("name") or "Pod"
    if pod.get("replicas"):
        title = f'{title} x{pod["replicas"]}'
    lines.append(trim_text(title, 16))
    spec_bits = []
    if pod.get("cpu"):
        spec_bits.append(f'{pod["cpu"]}C')
    if pod.get("memory"):
        spec_bits.append(f'{pod["memory"]}G')
    if spec_bits:
        lines.append(trim_text(" ".join(spec_bits), 16))
    elif pod.get("external_port"):
        lines.append(trim_text(str(pod["external_port"]), 16))
    return lines[:2]


def service_box_paragraphs(family: dict, title_color: str, body_color: str, include_ports: bool = True) -> list[dict]:
    lines = [{"text": trim_text(family["display_name"], 14), "size": 820, "color": title_color, "bold": True}]
    ports = compact_port_labels(family, limit=2)
    if include_ports and ports:
        lines.append({"text": f'端口 {trim_text(",".join(ports), 16)}', "size": 640, "color": body_color})
    if family.get("raw_endpoints"):
        lines.append({"text": trim_text(family["raw_endpoints"][0], 24), "size": 560, "color": body_color})
    if family.get("notes"):
        lines.append({"text": trim_text(family["notes"][0], 24), "size": 540, "color": body_color})
    return lines[:4]


def pod_paragraphs(pod: dict) -> list[dict]:
    title = pod["name"]
    if pod.get("replicas"):
        title = f'{title} x{pod["replicas"]}'
    port_text = ""
    if pod.get("container_port") and pod.get("external_port"):
        port_text = f'{pod["container_port"]} -> {pod["external_port"]}'
    elif pod.get("external_port"):
        port_text = pod["external_port"]

    lines = [{"text": trim_text(title, 16), "size": 780, "color": "1F4E79", "bold": True}]
    detail = []
    if pod.get("description"):
        detail.append(trim_text(pod["description"], 10))
    spec_bits = []
    if pod.get("cpu"):
        spec_bits.append(f'{pod["cpu"]}C')
    if pod.get("memory"):
        spec_bits.append(f'{pod["memory"]}G')
    if spec_bits:
        detail.append(" ".join(spec_bits))
    if detail:
        lines.append({"text": trim_text(" ".join(detail), 18), "size": 600, "color": "415A77"})
    if port_text:
        lines.append({"text": trim_text(port_text, 18), "size": 600, "color": "415A77"})
    mount = short_mount_path(pod.get("host_path") or pod.get("container_path"))
    if mount:
        lines.append({"text": trim_text(mount, 18), "size": 540, "color": "6C757D"})
    return lines


def worker_paragraphs(resource: dict) -> list[dict]:
    return server_box_paragraphs(resource, [], "FFFFFF", "EAF4FF", small=True, include_ports=False)


def worker_ip_only_paragraphs(resource: dict) -> list[dict]:
    return [{"text": compact_ip(resource.get("ip") or "-"), "size": 560, "color": "FFFFFF", "bold": True}]


def group_and_order_families(families: dict[str, dict]) -> dict[str, list[dict]]:
    order_map = {
        "access": ["lb", "nginx"],
        "application": ["gpaas", "preview", "k8s"],
        "data": ["pg", "mdd"],
        "platform": ["redis", "zookeeper", "mq", "elk", "nfs", "appstore"],
    }
    grouped = defaultdict(list)
    for family in families.values():
        grouped[family["zone"]].append(family)

    ordered = {}
    for zone, items in grouped.items():
        preferred = order_map.get(zone, [])
        ordered_items = []
        remaining = {item["key"]: item for item in items}
        for key in preferred:
            if key in remaining:
                ordered_items.append(remaining.pop(key))
        ordered_items.extend(sorted(remaining.values(), key=lambda item: item["display_name"]))
        ordered[zone] = ordered_items
    return ordered


def add_zone(builder: SlideBuilder, x: int, y: int, w: int, h: int, label: str, fill: str, border: str) -> None:
    builder.add_round_rect(
        x,
        y,
        w,
        h,
        "FFFFFF",
        ZONE_BORDER,
        paragraphs=None,
        border_width=ZONE_LINE,
        name=f"{label}Zone",
        shape="rect",
        dash="dash",
    )
    builder.add_text_box(
        x + builder.sx(120000),
        y - builder.sy(180000),
        builder.sx(1600000),
        builder.sy(260000),
        [{"text": label, "size": 1250, "color": TITLE_TEXT, "bold": True, "align": "l"}],
        name=f"{label}Label",
    )


def draw_server_group(
    builder: SlideBuilder,
    family: dict,
    rect: tuple[int, int, int, int],
    shell_fill: str,
    shell_border: str,
    server_fill: str,
    title_color: str,
    body_color: str,
    max_servers: int = 3,
    include_ports_in_boxes: bool = True,
    layout: str = "horizontal",
    shell_dash: str | None = "dash",
    shell_icon_target: str | None = None,
) -> tuple[int, int, int, int]:
    x, y, w, h = rect
    icon_target = ICON_TARGETS.get(FAMILY_ICON_KEYS.get(family["key"], "k8s"), ICON_TARGETS["k8s"])
    compact_shell = bool(family.get("compact_shell"))
    title_h = builder.sy(160000 if compact_shell else 220000)
    top_gap = builder.sy(20000 if compact_shell else 40000)
    side_gap = int(family.get("side_gap_x", builder.sx(50000 if compact_shell else 70000)))
    content_gap = builder.sy(30000 if compact_shell else 90000)
    bottom_pad = builder.sy(90000 if compact_shell else 160000)
    builder.add_round_rect(
        x,
        y,
        w,
        h,
        None,
        ZONE_BORDER,
        paragraphs=None,
        border_width=THIN_LINE,
        name=f'{family["key"]}Shell',
        shape="rect",
        dash=shell_dash,
    )
    if shell_icon_target:
        shell_icon_size = builder.cm(STANDARD_ICON_CM)
        if shell_icon_target in {ICON_TARGETS["mobile"], ICON_TARGETS["pc"], ICON_TARGETS["k8s"]}:
            shell_icon_size = builder.cm(SMALL_ICON_CM)
        builder.add_image(
            x + builder.sx(50000),
            y + builder.sy(10000),
            shell_icon_size,
            shell_icon_size,
            shell_icon_target,
            name=f'{family["key"]}ShellIcon',
        )
    builder.add_text_box(
        x + (builder.sx(220000) if shell_icon_target else 0),
        y + top_gap,
        w - (builder.sx(220000) if shell_icon_target else 0),
        title_h,
        [{"text": trim_text(family["display_name"], 18), "size": 880, "color": TITLE_TEXT, "bold": True, "align": "l"}],
        name=f'{family["key"]}Title',
    )

    resources = family.get("resources", [])[:max_servers]
    if resources:
        inner_x = x + side_gap
        inner_y = y + title_h + content_gap
        inner_w = w - side_gap * 2
        inner_h = h - title_h - bottom_pad
        fixed_card_h = family.get("fixed_server_card_h")
        if layout == "vertical":
            rows = min(len(resources), max_servers)
            gap = builder.sy(60000)
            if fixed_card_h:
                card_h = min(fixed_card_h, max(builder.sy(220000), inner_h - gap * max(rows - 1, 0)))
            else:
                card_h = (inner_h - gap * max(rows - 1, 0)) // max(rows, 1)
            card_w = inner_w
            for idx, resource in enumerate(resources):
                ry = inner_y + idx * (card_h + gap)
                builder.add_icon_label(
                    inner_x,
                    ry,
                    card_w,
                    card_h,
                    resource_icon_target(resource, family["key"]),
                    icon_detail_lines(resource),
                    name=f'{family["key"]}Server{idx+1}',
                    text_color=BODY_TEXT,
                    icon_cap=builder.sy(300000),
                )
        else:
            cols = min(len(resources), 3)
            gap = int(family.get("card_gap_x", builder.sx(70000)))
            card_w = (inner_w - gap * max(cols - 1, 0)) // max(cols, 1)
            card_h = min(fixed_card_h, inner_h) if fixed_card_h else inner_h
            for idx, resource in enumerate(resources):
                rx = inner_x + idx * (card_w + gap)
                builder.add_icon_label(
                    rx,
                    inner_y,
                    card_w,
                    card_h,
                    resource_icon_target(resource, family["key"]),
                    icon_detail_lines(resource),
                    name=f'{family["key"]}Server{idx+1}',
                    text_color=BODY_TEXT,
                    icon_cap=builder.sy(300000),
                )
    else:
        builder.add_icon_label(
            x + side_gap,
            y + title_h + content_gap,
            w - side_gap * 2,
            h - title_h - (bottom_pad + builder.sy(20000)),
            icon_target,
            [family["display_name"]] + compact_port_labels(family, limit=1),
            name=f'{family["key"]}ServiceOnly',
            text_color=BODY_TEXT,
            icon_cap=builder.sy(320000),
        )
    return rect


def render_diagram(title: str, families: dict[str, dict], pods: list[dict], slide_width: int, slide_height: int) -> tuple[str, list[dict], list[tuple[str, str]]]:
    builder = SlideBuilder(slide_width, slide_height)
    ordered = group_and_order_families(families)
    connections = []
    shift_x = builder.sx(350000)
    fixed_card_h = builder.cm(1.28)
    featured_card_h = builder.cm(1.2)

    builder.add_text_box(
        builder.sx(280000),
        builder.sy(120000),
        builder.sx(11600000),
        builder.sy(360000),
        [{"text": title, "size": 2400, "color": "214D8A", "bold": True}],
        name="Title",
    )

    access_box = (builder.sx(420000) + shift_x, builder.sy(1120000), builder.sx(2200000), builder.sy(2500000))
    app_box = (builder.sx(2850000) + shift_x, builder.sy(1120000), builder.sx(4650000), builder.sy(3100000))
    data_box = (builder.sx(7750000) + shift_x, builder.sy(1120000), builder.sx(3400000), builder.sy(3100000))
    middleware_cell_w = builder.cm(6.5)
    middleware_gap_x = builder.cm(0.08)
    middleware_gap_y = builder.cm(0.08)
    middleware_side_pad = builder.cm(0.3)

    access_families = {family["key"]: family for family in ordered.get("access", [])}
    lb = access_families.get("lb")
    nginx = access_families.get("nginx")
    external_lbs, internal_lb = split_lb_roles(lb)

    lb_rect = None
    mobile_lb_rect = None
    pc_lb_rect = None
    internal_lb_rect = None
    nginx_rect = None
    has_multi_lb_layout = bool(lb and external_lbs)
    if has_multi_lb_layout:
        mobile_lb_rect = (builder.sx(1100000) + shift_x, builder.sy(1820000), builder.sx(720000), builder.sy(640000))
        pc_lb_rect = (builder.sx(1100000) + shift_x, builder.sy(2520000), builder.sx(720000), builder.sy(640000))
        internal_lb_rect = (builder.sx(2100000) + shift_x, builder.sy(2180000), builder.sx(760000), builder.sy(720000))
        mobile_lb = {"key": "lb", "display_name": external_lbs[0].get("name") or "移动LB", "resources": [external_lbs[0]], "fixed_server_card_h": builder.cm(1.45), "compact_shell": True}
        draw_server_group(builder, mobile_lb, mobile_lb_rect, "FFE5BF", "FF8C00", "FF8C00", "FFFFFF", "FFF2DB", max_servers=1)
        if len(external_lbs) > 1:
            pc_lb = {"key": "lb", "display_name": external_lbs[1].get("name") or "PC端LB", "resources": [external_lbs[1]], "fixed_server_card_h": builder.cm(1.45), "compact_shell": True}
            draw_server_group(builder, pc_lb, pc_lb_rect, "FFE5BF", "FF8C00", "FF8C00", "FFFFFF", "FFF2DB", max_servers=1)
        if internal_lb:
            inner_family = {"key": "lb", "display_name": internal_lb.get("name") or "内部LB", "resources": [internal_lb], "fixed_server_card_h": builder.cm(1.45), "compact_shell": True}
            draw_server_group(builder, inner_family, internal_lb_rect, "FFE5BF", "FF8C00", "FF8C00", "FFFFFF", "FFF2DB", max_servers=1)
        nginx_rect = (builder.sx(2500000) + shift_x, builder.sy(2180000), builder.sx(760000), builder.sy(1180000))
        if nginx:
            draw_server_group(
                builder,
                nginx,
                nginx_rect,
                "D7EEF9",
                "006699",
                "006699",
                "FFFFFF",
                "D8F3FF",
                max_servers=2,
                layout="vertical",
            )

    if has_multi_lb_layout:
        user_group_x = builder.sx(260000)
        user_group_y = builder.sy(2160000)
        user_group_w = builder.sx(420000)
        user_group_h = builder.sy(860000)
        mobile_user_y = user_group_y + builder.sy(70000)
        pc_user_y = user_group_y + builder.sy(390000)
        builder.add_round_rect(user_group_x, user_group_y, user_group_w, user_group_h, None, ZONE_BORDER, paragraphs=None, border_width=THIN_LINE, name="UserGroup", shape="rect", dash="dash")
        builder.add_image(user_group_x + builder.sx(90000), mobile_user_y, builder.cm(SMALL_ICON_CM), builder.cm(SMALL_ICON_CM), ICON_TARGETS["mobile"], name="MobileIcon")
        builder.add_image(user_group_x + builder.sx(90000), pc_user_y, builder.cm(SMALL_ICON_CM), builder.cm(SMALL_ICON_CM), ICON_TARGETS["pc"], name="PCIcon")
        builder.add_text_box(user_group_x + builder.sx(10000), mobile_user_y + builder.cm(SMALL_ICON_CM) + builder.sy(10000), builder.sx(360000), builder.sy(120000), [{"text": "Mobile", "size": STANDARD_ICON_TEXT_SIZE, "color": BODY_TEXT, "bold": True}], name="MobileText", wrap="none")
        builder.add_text_box(user_group_x + builder.sx(10000), pc_user_y + builder.cm(SMALL_ICON_CM) + builder.sy(10000), builder.sx(360000), builder.sy(120000), [{"text": "PC", "size": STANDARD_ICON_TEXT_SIZE, "color": BODY_TEXT, "bold": True}], name="PCText", wrap="none")
        builder.add_connector(user_group_x + user_group_w, mobile_user_y + builder.sy(100000), mobile_lb_rect[0], mobile_lb_rect[1] + builder.sy(160000), "FF8C00")
        if len(external_lbs) > 1:
            builder.add_connector(user_group_x + user_group_w, pc_user_y + builder.sy(100000), pc_lb_rect[0], pc_lb_rect[1] + builder.sy(160000), "FF8C00")
        if internal_lb:
            builder.add_connector(mobile_lb_rect[0] + mobile_lb_rect[2], mobile_lb_rect[1] + builder.sy(160000), internal_lb_rect[0], internal_lb_rect[1] + builder.sy(160000), "FF8C00")
            if len(external_lbs) > 1:
                builder.add_connector(pc_lb_rect[0] + pc_lb_rect[2], pc_lb_rect[1] + builder.sy(160000), internal_lb_rect[0], internal_lb_rect[1] + builder.sy(460000), "FF8C00")
    else:
        user_group_x = builder.sx(280000)
        user_group_y = builder.sy(2436428)
        builder.add_image(user_group_x + builder.sx(90000), user_group_y + builder.sy(70000), builder.cm(SMALL_ICON_CM), builder.cm(SMALL_ICON_CM), ICON_TARGETS["mobile"], name="MobileIcon")
        builder.add_image(user_group_x + builder.sx(90000), user_group_y + builder.sy(390000), builder.cm(SMALL_ICON_CM), builder.cm(SMALL_ICON_CM), ICON_TARGETS["pc"], name="PCIcon")
        builder.add_round_rect(user_group_x, user_group_y, builder.sx(420000), builder.sy(860000), None, ZONE_BORDER, paragraphs=None, border_width=THIN_LINE, name="UserGroup", shape="rect", dash="dash")
        builder.add_text_box(user_group_x + builder.sx(10000), user_group_y + builder.sy(70000) + builder.cm(SMALL_ICON_CM) + builder.sy(10000), builder.sx(360000), builder.sy(120000), [{"text": "Mobile", "size": STANDARD_ICON_TEXT_SIZE, "color": BODY_TEXT, "bold": True}], name="MobileText", wrap="none")
        builder.add_text_box(user_group_x + builder.sx(10000), user_group_y + builder.sy(390000) + builder.cm(SMALL_ICON_CM) + builder.sy(10000), builder.sx(360000), builder.sy(120000), [{"text": "PC", "size": STANDARD_ICON_TEXT_SIZE, "color": BODY_TEXT, "bold": True}], name="PCText", wrap="none")
        # Compact single-LB layout for sheets without distinct mobile/pc/internal LB resources.
        lb_rect = (builder.sx(780000) + shift_x, builder.sy(2344028), builder.sx(640000), builder.sy(900000))
        if lb:
            draw_server_group(builder, lb, lb_rect, "FFE5BF", "FF8C00", "FF8C00", "FFFFFF", "FFF2DB", max_servers=1)
        nginx_rect = (builder.sx(1680000) + shift_x, builder.sy(2225000), builder.sx(900000), builder.sy(1180000))
        if nginx:
            draw_server_group(
                builder,
                nginx,
                nginx_rect,
                "D7EEF9",
                "006699",
                "006699",
                "FFFFFF",
                "D8F3FF",
                max_servers=2,
                layout="vertical",
            )
        access_chain_y = builder.sy(2785000)
        user_anchor_x = user_group_x + builder.sx(420000)
        user_anchor_y = access_chain_y
        builder.add_connector(user_anchor_x, user_anchor_y, lb_rect[0], access_chain_y, "FF8C00")
        if lb and nginx:
            builder.add_connector(lb_rect[0] + lb_rect[2], access_chain_y, nginx_rect[0], access_chain_y, "FF8C00")
    if lb and lb.get("ports"):
        label_target_rect = mobile_lb_rect if has_multi_lb_layout and mobile_lb_rect else lb_rect
        builder.add_line_label(
            (builder.sx(700000) if has_multi_lb_layout else builder.sx(980000)),
            (builder.sy(2460000) if has_multi_lb_layout else lb_rect[1] + lb_rect[3] // 2),
            label_target_rect[0],
            label_target_rect[1] + label_target_rect[3] // 2,
            ",".join(lb["ports"][:2]),
            "A23B1E",
            w=builder.sx(360000),
            y_offset=-builder.sy(120000),
        )
        connections.append({"from": "用户", "to": lb["display_name"], "ports": lb["ports"][:2]})

    app_families = {family["key"]: family for family in ordered.get("application", [])}
    gpaas = app_families.get("gpaas")
    preview = app_families.get("preview")
    k8s = app_families.get("k8s")

    gpaas_rect = (builder.sx(3273000) + shift_x, builder.sy(1122594), builder.sx(1750000), builder.sy(760000))
    if gpaas:
        gpaas["fixed_server_card_h"] = builder.cm(1.38)
        gpaas["compact_shell"] = True
        draw_server_group(builder, gpaas, gpaas_rect, "DDF5E6", "2E8B57", "2E8B57", "FFFFFF", "EAF9F0", max_servers=1, shell_dash=None, shell_icon_target=ICON_TARGETS["k8s"])

    preview_rect = (builder.sx(5480000) + shift_x, builder.sy(1190731), 1096000, builder.sy(760000))
    if preview:
        preview["fixed_server_card_h"] = builder.cm(1.18)
        preview["compact_shell"] = True
        draw_server_group(builder, preview, preview_rect, "E4F2EC", "5B8E7D", "006699", "FFFFFF", "D8F3FF", max_servers=1)

    cluster_rect = (builder.sx(3000000) + shift_x, builder.sy(2100000), builder.sx(4400000), builder.sy(1680000))
    if k8s:
        builder.add_round_rect(*cluster_rect, "FFFFFF", ZONE_BORDER, paragraphs=None, border_width=ZONE_LINE, name="K8SOuter", shape="rect")
        builder.add_image(cluster_rect[0] + builder.sx(110000), cluster_rect[1] + builder.sy(70000), builder.cm(SMALL_ICON_CM), builder.cm(SMALL_ICON_CM), "../media/k8s-icon.png", name="K8SIcon")
        header_lines = [{"text": f'K8S容器集群 ({len(k8s.get("resources", [])) or "?"} 节点)', "size": 1080, "color": TITLE_TEXT, "bold": True}]
        if k8s.get("ports"):
            header_lines.append({"text": f'对外端口 {trim_text(",".join(k8s["ports"][:4]), 22)}', "size": 680, "color": BODY_TEXT})
        vip_note = ""
        for service in k8s.get("services", []):
            if service.get("vip_hint"):
                vip_note = service["vip_hint"]
                break
        if vip_note:
            header_lines.append({"text": f'API VIP {trim_text(vip_note, 18)}:6443', "size": 640, "color": BODY_TEXT})
        builder.add_text_box(
            cluster_rect[0] + builder.sx(380000),
            cluster_rect[1] + builder.sy(80000),
            cluster_rect[2] - builder.sx(500000),
            builder.sy(360000),
            header_lines,
            name="K8SHeader",
        )

        workers = k8s.get("resources", [])
        worker_gap = builder.sx(100000)
        worker_area_x = cluster_rect[0] + builder.sx(140000)
        worker_area_w = cluster_rect[2] - builder.sx(280000)
        if pods:
            pod_area_x = cluster_rect[0] + builder.sx(140000)
            pod_area_y = cluster_rect[1] + builder.sy(430000)
            pod_area_w = cluster_rect[2] - builder.sx(280000)
            pod_area_h = builder.sy(700000)
            builder.add_round_rect(
                pod_area_x,
                pod_area_y,
                pod_area_w,
                pod_area_h,
                "BFDFFF",
                None,
                paragraphs=None,
                border_width=0,
                name="PodAreaBg",
                shape="rect",
            )
            pod_gap_x = builder.sx(90000)
            pod_gap_y = builder.sy(90000)
            max_pod_rows = max(1, (pod_area_h + pod_gap_y) // (fixed_card_h + pod_gap_y))
            columns = max(1, math.ceil(len(pods) / max_pod_rows))
            pod_w = (pod_area_w - pod_gap_x * (columns - 1)) // max(columns, 1)
            pod_h = fixed_card_h
            for index, pod in enumerate(pods):
                col = index % columns
                row = index // columns
                pod_x = pod_area_x + col * (pod_w + pod_gap_x)
                pod_y = pod_area_y + row * (pod_h + pod_gap_y)
                builder.add_icon_label(
                    pod_x,
                    pod_y + builder.sy(10000),
                    pod_w,
                    pod_h - builder.sy(20000),
                    ICON_TARGETS["pod"],
                    pod_icon_lines(pod),
                    name=f"Pod{index+1}",
                    text_color=BODY_TEXT,
                    icon_cap=builder.sy(210000),
                    text_size=460,
                )
            worker_y = pod_area_y + pod_area_h + builder.sy(90000)
        else:
            # No pods: give the full content area back to the node grid.
            worker_y = cluster_rect[1] + builder.sy(470000)

        worker_area_h = max(builder.sy(240000), cluster_rect[1] + cluster_rect[3] - worker_y - builder.sy(110000))
        if len(workers) > 4:
            worker_cols = min(4, len(workers))
            worker_rows = math.ceil(len(workers) / worker_cols)
            worker_gap_y = builder.sy(40000)
            worker_w = (worker_area_w - worker_gap * max(0, worker_cols - 1)) // worker_cols
            available_h = max(builder.sy(180000), worker_area_h - worker_gap_y * max(worker_rows - 1, 0))
            worker_h = max(builder.sy(180000), available_h // max(worker_rows, 1))
            for idx, worker in enumerate(workers):
                col = idx % worker_cols
                row = idx // worker_cols
                worker_x = worker_area_x + col * (worker_w + worker_gap)
                worker_box_y = worker_y + row * (worker_h + worker_gap_y)
                builder.add_icon_label(worker_x, worker_box_y, worker_w, worker_h, ICON_TARGETS["k8s_node"], icon_detail_lines(worker, only_ip=True), name=f"Worker{idx+1}", text_color=BODY_TEXT, icon_cap=builder.sy(190000), text_size=460)
        else:
            worker_w = (worker_area_w - worker_gap * max(0, len(workers) - 1)) // max(len(workers), 1)
            worker_h = min(builder.sy(620000), worker_area_h)
            for idx, worker in enumerate(workers[:4]):
                worker_x = worker_area_x + idx * (worker_w + worker_gap)
                builder.add_icon_label(worker_x, worker_y, worker_w, worker_h, ICON_TARGETS["k8s_node"], icon_detail_lines(worker), name=f"Worker{idx+1}", text_color=BODY_TEXT, icon_cap=builder.sy(210000), text_size=480)

    if internal_lb and k8s:
        builder.add_connector(
            internal_lb_rect[0] + internal_lb_rect[2],
            internal_lb_rect[1] + builder.sy(310000),
            cluster_rect[0],
            internal_lb_rect[1] + builder.sy(310000),
            "FF8C00",
        )
        connections.append({"from": internal_lb.get("name") or "内部LB", "to": k8s["display_name"], "ports": []})
    elif lb and nginx and has_multi_lb_layout:
        source_rect = internal_lb_rect if internal_lb else lb_rect
        builder.add_connector(source_rect[0] + source_rect[2], source_rect[1] + builder.sy(310000), nginx_rect[0], source_rect[1] + builder.sy(310000), "FF8C00")
        connections.append({"from": lb["display_name"], "to": nginx["display_name"], "ports": lb.get("ports", [])[:2]})

    if nginx and k8s:
        start_x = nginx_rect[0] + nginx_rect[2]
        start_y = internal_lb_rect[1] + builder.sy(310000) if internal_lb else builder.sy(2785000)
        end_x = cluster_rect[0]
        end_y = start_y
        builder.add_connector(start_x, start_y, end_x, end_y, "214D8A")
        if k8s.get("ports"):
            builder.add_line_label(start_x, start_y, end_x, end_y, ",".join(k8s["ports"][:3]), "214D8A", w=builder.sx(420000), h=builder.sy(130000))
        connections.append({"from": nginx["display_name"], "to": k8s["display_name"], "ports": k8s.get("ports", [])[:3]})

    if gpaas and k8s:
        gpaas_line_x = gpaas_rect[0] + gpaas_rect[2] // 2
        builder.add_connector(
            gpaas_line_x,
            gpaas_rect[1] + gpaas_rect[3],
            gpaas_line_x,
            cluster_rect[1],
            "2E8B57",
            name="GpaasToK8S",
        )
        connections.append({"from": gpaas["display_name"], "to": k8s["display_name"], "ports": ["6443"]})

    if preview and k8s:
        preview_line_x = preview_rect[0] + preview_rect[2] // 2
        cluster_line_x = min(max(preview_line_x, cluster_rect[0] + builder.sx(300000)), cluster_rect[0] + cluster_rect[2] - builder.sx(300000))
        builder.add_connector(cluster_line_x, cluster_rect[1], preview_line_x, preview_rect[1] + preview_rect[3], "5B8E7D")
        connections.append({"from": k8s["display_name"], "to": preview["display_name"], "ports": []})

    data_families = ordered.get("data", [])
    data_rects = []
    if data_families:
        gap = builder.sy(160000)
        box_h = min(builder.sy(1080000), (data_box[3] - builder.sy(500000) - gap * (len(data_families) - 1)) // len(data_families))
        for idx, family in enumerate(data_families[:3]):
            rect = (
                data_box[0] + builder.sx(180000),
                data_box[1] + builder.sy(360000) + idx * (box_h + gap),
                data_box[2] - builder.sx(360000),
                box_h,
            )
            data_rects.append((family, rect))
            draw_server_group(builder, family, rect, "F6E5DE", "A23B1E", "FCEEE8", "7A2E17", "5C4033", max_servers=3)
            if k8s:
                start_x = cluster_rect[0] + cluster_rect[2]
                start_y = cluster_rect[1] + cluster_rect[3] // 2
                end_x = rect[0]
                end_y = rect[1] + rect[3] // 2
                builder.add_connector(start_x, start_y, end_x, end_y, "A23B1E")
                if family.get("ports"):
                    builder.add_line_label(start_x, start_y, end_x, end_y, ",".join(family["ports"][:2]), "A23B1E", w=builder.sx(360000), h=builder.sy(130000))
                connections.append({"from": k8s["display_name"], "to": family["display_name"], "ports": family.get("ports", [])[:2]})

    nfs = families.get("nfs")
    nfs_rect = None
    if nfs:
        nfs["compact_shell"] = True
        nfs_rect = (
            cluster_rect[0] + builder.sx(150000),
            cluster_rect[1] + cluster_rect[3] + builder.sy(130000),
            builder.sx(1180000),
            builder.sy(650000),
        ) if k8s else (
            builder.sx(5000000),
            builder.sy(3950000),
            builder.sx(1180000),
            builder.sy(650000),
        )
        draw_server_group(builder, nfs, nfs_rect, "FFF7D6", "C99700", "FFFFFF", "6B5B00", "6B5B00", max_servers=1)
        if k8s:
            line_x = nfs_rect[0] + nfs_rect[2] // 2
            builder.add_connector(line_x, cluster_rect[1] + cluster_rect[3], line_x, nfs_rect[1], "C99700", name="K8SToNFS")
            connections.append({"from": k8s["display_name"], "to": nfs["display_name"], "ports": nfs.get("ports", [])[:2]})

    bottom_box = (
        (nfs_rect[0] if nfs_rect else builder.sx(4450000) + shift_x),
        (nfs_rect[1] + nfs_rect[3] + builder.sy(240000) if nfs_rect else builder.sy(4480000)),
        middleware_side_pad * 2 + middleware_cell_w * 2 + middleware_gap_x,
        builder.sy(1600000),
    )
    add_zone(builder, *bottom_box, "中间件", "F6F6F6", "6C757D")

    platform_families = [family for family in ordered.get("platform", []) if family["key"] in {"redis", "zookeeper", "mq", "elk"}]
    if platform_families:
        cols = 2 if len(platform_families) <= 4 else 3
        rows_needed = math.ceil(len(platform_families) / cols)
        gap_x = middleware_gap_x
        gap_y = middleware_gap_y
        cell_w = middleware_cell_w
        cell_h = (bottom_box[3] - builder.sy(220000) - gap_y * (rows_needed - 1)) // rows_needed

        color_map = {
            "redis": ("E8F0FF", "3D5A80", "293241"),
            "zookeeper": ("E8F0FF", "3D5A80", "293241"),
            "mq": ("FFF1E6", "D97706", "7C2D12"),
            "elk": ("F3F0FF", "4361EE", "2B2D42"),
        }

        for idx, family in enumerate(platform_families):
            col = idx % cols
            row = idx // cols
            family["fixed_server_card_h"] = fixed_card_h
            family["compact_shell"] = True
            family["side_gap_x"] = middleware_gap_x
            family["card_gap_x"] = middleware_gap_x
            rect = (
                bottom_box[0] + middleware_side_pad + col * (cell_w + gap_x),
                bottom_box[1] + builder.sy(100000) + row * (cell_h + gap_y),
                cell_w,
                cell_h,
            )
            fill, border, title_color = color_map.get(family["key"], ("F1F3F5", "6C757D", "495057"))
            suppress_ports = family["key"] in {"redis", "zookeeper", "mq", "elk"}
            draw_server_group(
                builder,
                family,
                rect,
                fill,
                border,
                "FFFFFF",
                title_color,
                "495057",
                max_servers=3,
                include_ports_in_boxes=not suppress_ports,
            )
            if k8s:
                start_x = cluster_rect[0] + cluster_rect[2] // 2
                start_y = cluster_rect[1] + cluster_rect[3]
                end_x = rect[0] + rect[2] // 2
                end_y = rect[1]
                builder.add_connector(start_x, start_y, end_x, end_y, border, width=THIN_LINE)
                if family.get("ports"):
                    builder.add_line_label(start_x, start_y, end_x, end_y, ",".join(family["ports"][:2]), border, w=builder.sx(340000), h=builder.sy(130000))
                connections.append({"from": k8s["display_name"], "to": family["display_name"], "ports": family.get("ports", [])[:3]})

    appstore = families.get("appstore")
    if appstore and nginx_rect:
        appstore_rect = (
            nginx_rect[0] + builder.sx(40000),
            nginx_rect[1] + nginx_rect[3] + builder.sy(140000),
            builder.sx(760000),
            builder.sy(460000),
        )
        appstore_lines = [appstore["display_name"]]
        appstore_lines.extend(compact_port_labels(appstore, limit=1))
        builder.add_icon_label(
            appstore_rect[0],
            appstore_rect[1],
            appstore_rect[2],
            appstore_rect[3],
            ICON_TARGETS["appstore"],
            appstore_lines,
            name="AppStoreStandalone",
            text_color=BODY_TEXT,
            text_size=STANDARD_ICON_TEXT_SIZE,
        )
        start_x = nginx_rect[0] + nginx_rect[2] // 2
        start_y = nginx_rect[1] + nginx_rect[3]
        end_x = appstore_rect[0] + appstore_rect[2] // 2
        end_y = appstore_rect[1]
        builder.add_connector(start_x, start_y, end_x, end_y, "2A9D8F", name="NginxToAppStore")
        if appstore.get("ports"):
            builder.add_line_label(start_x, start_y, end_x, end_y, ",".join(appstore["ports"][:1]), "2A9D8F", w=builder.sx(300000), h=builder.sy(120000), y_offset=-builder.sy(160000))
        connections.append({"from": nginx["display_name"] if nginx else "Nginx", "to": appstore["display_name"], "ports": appstore.get("ports", [])[:1]})

    return builder.build(), connections, builder.image_rels


def build_slide_rels_xml(existing_rels_xml: bytes, image_rels: list[tuple[str, str]]) -> bytes:
    root = ET.fromstring(existing_rels_xml)
    ns_uri = "http://schemas.openxmlformats.org/package/2006/relationships"
    for rel_id, target in image_rels:
        ET.SubElement(
            root,
            f"{{{ns_uri}}}Relationship",
            {
                "Id": rel_id,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                "Target": target,
            },
        )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def write_pptx_from_template(template_path: Path, output_path: Path, slide_xml: str, image_rels: list[tuple[str, str]] | None = None) -> None:
    image_rels = image_rels or []
    extra_media = {}
    if image_rels:
        for _, target in image_rels:
            filename = "ppt/media/" + Path(target).name
            source_path = ICON_TARGET_TO_SOURCE.get(target)
            if source_path and source_path.exists():
                extra_media[filename] = source_path.read_bytes()
    tmp_output = output_path.with_suffix(output_path.suffix + ".tmp")
    with zipfile.ZipFile(template_path, "r") as source, zipfile.ZipFile(tmp_output, "w", compression=zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            if info.filename == "ppt/slides/slide1.xml":
                data = slide_xml.encode("utf-8")
            elif info.filename == "ppt/slides/_rels/slide1.xml.rels" and image_rels:
                data = build_slide_rels_xml(source.read(info.filename), image_rels)
            else:
                data = source.read(info.filename)
            new_info = zipfile.ZipInfo(info.filename, info.date_time)
            new_info.compress_type = zipfile.ZIP_DEFLATED
            new_info.comment = info.comment
            new_info.extra = info.extra
            new_info.create_system = info.create_system
            new_info.external_attr = info.external_attr
            new_info.internal_attr = info.internal_attr
            target.writestr(new_info, data)
        for filename, data in extra_media.items():
            target.writestr(filename, data)
    tmp_output.replace(output_path)


def validate_pptx(path: Path) -> dict:
    with zipfile.ZipFile(path) as archive:
        broken = archive.testzip()
        slide_xml = archive.read("ppt/slides/slide1.xml")
        ET.fromstring(slide_xml)
    return {"testzip": broken or "", "slide_xml_valid": True}


def build_summary(
    workbook_path: Path,
    output_pptx: Path,
    output_json: Path,
    title: str,
    resources: list[dict],
    services: list[dict],
    pods: list[dict],
    families: dict[str, dict],
    unmatched_resources: list[dict],
    conflicts: list[dict],
    connections: list[dict],
    validation: dict,
    upload_result: dict | None,
) -> dict:
    return {
        "workbook": str(workbook_path),
        "title": title,
        "output_pptx": str(output_pptx),
        "output_json": str(output_json),
        "counts": {
            "resources": len(resources),
            "services": len(services),
            "pods": len(pods),
            "families": len(families),
        },
        "families": {
            key: {
                "display_name": family["display_name"],
                "zone": family["zone"],
                "ports": family["ports"],
                "port_labels": family["port_labels"],
                "resource_count": len(family["resources"]),
                "resources": family["resources"],
                "notes": family["notes"],
            }
            for key, family in families.items()
        },
        "connections": connections,
        "warnings": {
            "port_conflicts": conflicts,
            "unmatched_resources": unmatched_resources,
        },
        "validation": validation,
        "upload": upload_result,
    }


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate a PowerPoint architecture diagram from an ops workbook.")
    parser.add_argument("--workbook", required=True, help="Path to the customer workbook (.xlsx)")
    parser.add_argument("--template", help="Path to the PPTX template", default=str(SCRIPT_DIR.parent / "references" / "arch-model.pptx"))
    parser.add_argument("--output-dir", help="Output directory", default=str(SCRIPT_DIR.parent / "outputs"))
    parser.add_argument("--deck-name", help="Output PPTX file name without extension")
    parser.add_argument("--title", help="Override slide title")
    parser.add_argument("--upload", action="store_true", help="Upload the generated deck after creation")
    parser.add_argument("--emit-summary-only", action="store_true", help="Skip upload and emit summary JSON path")
    return parser


def main() -> int:
    parser = build_arg_parser()
    args = parser.parse_args()

    workbook_path = Path(args.workbook).expanduser().resolve()
    template_path = Path(args.template).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    reader = XlsxReader(workbook_path)
    try:
        roles = identify_sheet_roles(reader)
        if "resource" not in roles:
            raise ValueError("此文档不符合要求，请提供带有ip地址列的表格。")

        resources = parse_resource_sheet(reader.read_sheet_matrix(roles["resource"]))
        pods = parse_pod_sheet(reader.read_sheet_matrix(roles["pod"])) if "pod" in roles else []
        services = parse_service_sheet(reader.read_sheet_matrix(roles["service"])) if "service" in roles else synthesize_services_from_resources(resources, pods)
    finally:
        reader.close()

    families, unmatched_resources = build_families(resources, services, pods)
    conflicts = detect_port_conflicts(services, pods)

    env_name = next((resource["env"] for resource in resources if resource.get("env")), "") or next((pod["env"] for pod in pods if pod.get("env")), "")
    if env_name and "环境" not in env_name:
        env_name = f"{env_name}环境"
    title = args.title or f'{env_name or "客户"}部署图'

    slide_width, slide_height = read_template_slide_size(template_path)
    slide_xml, connections, image_rels = render_diagram(title, families, pods, slide_width, slide_height)

    deck_stem = args.deck_name or f"{workbook_path.stem}-architecture"
    output_pptx = output_dir / f"{deck_stem}.pptx"
    output_json = output_dir / f"{deck_stem}.json"

    write_pptx_from_template(template_path, output_pptx, slide_xml, image_rels)
    validation = validate_pptx(output_pptx)

    upload_result = None
    if args.upload and not args.emit_summary_only:
        upload_result = upload_file(file_path=output_pptx, **load_upload_config())

    summary = build_summary(
        workbook_path,
        output_pptx,
        output_json,
        title,
        resources,
        services,
        pods,
        families,
        unmatched_resources,
        conflicts,
        connections,
        validation,
        upload_result,
    )

    output_json.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"pptx": str(output_pptx), "summary": str(output_json), "upload": upload_result}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
