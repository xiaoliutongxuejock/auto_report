# Auto_Report — 自动化日报推送系统

基于 Python 的自动化数据报表生成与推送工具，每日自动采集 FileServer 空间数据、设备效能统计及机房温湿度数据，生成截图并通过企业微信机器人推送到群聊。

## 项目结构

```
Auto_Report/
├── FileServer_Report.py      # FileServer 空间增长报表
├── Performance_Report.py     # 效能统计 & 温湿度报表
├── README.md
└── ...
```

## 功能概览

| 脚本 | 功能 | 数据源 | 推送内容 |
|------|------|--------|----------|
| `FileServer_Report.py` | FileServer 剩余空间及使用率日报 | 每日 CSV 文件 | 文本 + 1 张截图 |
| `Performance_Report.py` | ACF/CF/Cell 效能 & 机房温湿度日报 | 多个 Excel 统计表 | 文本 + 4 张截图 |

---

## FileServer_Report.py

### 工作流程

1. **运行外部脚本** — 提权运行批处理脚本获取最新 FileServer 空间数据
2. **查找当天数据文件** — 在指定目录匹配当日 `fileserver_temp_YYYYMMDD.csv`
3. **数据导入** — 读取 CSV 最后一行，写入 `fileserver.xlsx` 末尾并复制上一行格式
4. **截图导出** — 通过 Excel COM 导出冻结区域的图表、标题行（18-19行）及最近 7 行数据，拼接为一张 PNG
5. **企业微信推送** — 发送文本消息 + 截图到企业微信群

### 核心依赖

- `win32com.client` — 操作 Excel COM 接口
- `PIL (Pillow)` — 图片裁剪与拼接
- `requests` — 企业微信 Webhook 推送
- `hashlib` / `base64` — 图片消息编码

### 关键配置

| 配置项 | 说明 |
|--------|------|
| `base_path` | 根目录路径 |
| `folder` | CSV 文件存放目录 |
| `target_file` | 目标 Excel 文件 (fileserver.xlsx) |
| `title_rows` | 需要截图导出的标题行号 |
| `WEBHOOK_URL` | 企业微信机器人 Webhook 地址 |

---

## Performance_Report.py

### 工作流程

1. **运行外部脚本** — 依次执行 4 个脚本：重连 FileServer、刷新凭证、生成 ACF/CF/Cell 效能统计、生成温湿度统计
2. **数据搬运** — 将 4 个源 Excel 的最后一行数据（含格式）追加到对应的目标 Excel 中：

   | 源文件 | 目标文件 |
   |--------|----------|
   | `acf效能统计表.xlsx` | `acftest效能统计.xlsx` |
   | `cell效能统计表.xlsx` | `celltest效能统计.xlsx` |
   | `cf效能统计表.xlsx` | `cftest效能统计.xlsx` |
   | `温湿度情况统计表.xlsx` | `机房温湿度统计.xlsx` |

3. **截图导出** — 对 4 个目标 Excel 分别导出图表 + 标题行 + 最后 7 行数据，保存为 4 张 PNG
4. **企业微信推送** — 发送文本消息 + 4 张截图到企业微信群

### 核心依赖

- `pandas` + `openpyxl` — 读取源 Excel 并写入目标 Excel（含格式复制）
- `win32com.client` — Excel COM 截图导出
- `PIL (Pillow)` — 图片裁剪与拼接
- `requests` — 企业微信 Webhook 推送

### 关键配置

| 配置项 | 说明 |
|--------|------|
| `BASE_DIR` | 根目录路径 |
| `EXCEL_PAIRS` | 源文件→目标文件的映射列表 |
| `FILES` | 截图配置：文件路径、Sheet名、输出路径、标题行 |
| `WEBHOOK_URL` | 企业微信机器人 Webhook 地址 |

---

## 环境要求

- **操作系统**: Windows（依赖 `win32com` 和外部 `.bat` 脚本）
- **Python**: 3.7+
- **Microsoft Excel**: 必须安装（COM 接口依赖）

### Python 依赖安装

```bash
pip install pywin32 pillow pandas openpyxl requests
```

---

## 使用方式

### 手动运行

```bash
# FileServer 空间报表
python FileServer_Report.py

# 效能 & 温湿度报表
python Performance_Report.py
```

### 定时任务（推荐）

通过 Windows 任务计划程序设置每日定时执行，例如：
- **FileServer_Report.py** — 每日 10:00
- **Performance_Report.py** — 每日 8:00 / 12:00 / 16:00

---

## 注意事项

1. **路径依赖** — 脚本中的文件路径基于 `D:\刘欢_勿删\auto_report`，迁移环境时需修改 `base_path` / `BASE_DIR` 配置
2. **Excel COM 限制** — 截图功能依赖 Excel COM 接口，运行时 Excel 窗口会自动最小化，**RDP 最小化不影响截图**
3. **企业微信限制** — 图片消息单张不能超过 2MB，脚本已做大小检查
4. **外部脚本依赖** — 数据生成依赖外部 `.bat` 脚本和 `.exe` 工具，需确保路径正确且可执行
5. **日志** — 所有运行日志保存在 `logs/` 目录下，便于排查问题

---

## 常见问题

| 问题 | 可能原因 | 解决方法 |
|------|----------|----------|
| 未找到当天 CSV 文件 | 外部脚本未正常生成数据 | 检查 `SCRIPT_1` / `SCRIPT_2` 是否执行成功 |
| Excel 操作失败 | Excel 进程残留 | 手动结束 Excel 进程后重试 |
| 图片推送失败 | 图片超过 2MB 或网络问题 | 检查图片大小及 Webhook 地址有效性 |
| 截图空白 | RDP 会话断开 | 本脚本采用 Chart 导出方式，不依赖屏幕截图，理论上不受影响 |
