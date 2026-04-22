# Workbook Format

The generator is built around workbooks shaped like `zyqd-test.xlsx`.

## Required Sheets

- `1-资源清单`
- `2-服务清单`
- `3-容器应用`

## 1-资源清单

Primary source for infrastructure inventory.

Important columns:

- `环境`
- `用途`
- `资源类型`
- `机器名/服务名`
- `操作系统/服务版本`
- `IP/域名`
- `虚拟IP`
- `LB-IP`
- `核心 vCPU`
- `内存（G)`
- `系统盘`
- `数据盘`
- `数据盘挂载目录`
- `类型`

Use this sheet to label machine names, IPs, VIPs, LB IPs, CPU, memory, and disk sizing.

## 2-服务清单

Primary source for service topology.

Important columns:

- `类别`
- `服务名称`
- `服务用途`
- `部署模式`
- `服务资源`
- `服务访问端口`
- `备注`

Use this sheet to group services into access, container, middleware, database, and other zones, and to decide access arrows and port labels.

## 3-容器应用

Primary source for pod-level enrichment.

Important columns:

- `容器节点名称`
- `服务描述`
- `POD资源配置`
- `副本数`
- `JVM堆内存配置(G)`
- `容器端口`
- `外部端口`
- `容器内路径`
- `宿主机路径(落盘在共享存储）`

Use this sheet to draw pod service blocks inside the K8s section.

## Parsing Notes

- Blank `环境` or `类别` cells are treated as inherited from the previous populated row.
- IP addresses may be partially masked with `XX`; matching logic falls back to the last numeric segment when needed.
- The generator ignores passwords.
- Extra sheets do not block generation.
