# Workbook Format

The skill is built around workbooks shaped like `zyqd-test.xlsx`.

## Required Resource Input

Accept only resource sheets that contain machine-level deployment facts.

Accepted patterns:

- Traditional resource inventory with fields such as `机器名/服务名`, `IP/域名`, `数据盘`
- Simple resource inventory with fields such as `资产名称`, `IP`, `CPU`, `内存`
- Standard server list with fields such as `环境`, `服务组名称`, `服务器名称`, `服务器IP`

Reject planning-only or total-only resource summary sheets. If a workbook only provides planned quantities, total CPU, or total memory without per-machine endpoints or names, it is unsupported for this skill.

## Service Input

Service sheets are optional but preferred.

Common fields:

- `类别`
- `服务名称`
- `服务用途`
- `部署模式`
- `服务资源`
- `服务访问端口`
- `备注`

Use service data to group zones, infer arrows, and label exposed ports.

## Container Or Pod Input

Container sheets are optional but should be used whenever present.

The sheet does not need a fixed tab index. It may be named `3-容器应用`, `容器配置`, or another tab name containing container or pod hints. Recognize it by content when the first rows contain most of these fields:

- `容器节点名称`
- `服务描述`
- `POD资源配置`
- `副本数`
- `小计资源`
- `JVM堆内存配置(G)`
- `容器端口`
- `外部端口`
- `容器内路径`
- `宿主机路径`

Use this sheet to render pod blocks inside the K8s region.

## Parsing Notes

- Blank `环境` or `类别` cells inherit the previous populated value.
- Pod sheet headers may span two rows; parse by header meaning, not fixed column offsets.
- Pod environments like `生产（namespace：xxx）` should collapse back to the matching resource environment such as `生产` when possible.
- IP addresses may be partially masked with `XX`; matching may fall back to the last numeric segment.
- Ignore passwords and credential columns.
- Extra sheets do not block generation.
