# 🖨️ Quick Print，自助打印与扫描系统（基于 Flask）

本项目由AI生成，提供一个运行于 **Windows 10** 上的 Web 服务，用户通过浏览器上传文件即可打印，同时支持扫描仪操作。支持打印图像、PDF、Word、Excel、TXT 文件，并可自定义打印机、扫描分辨率、颜色模式等选项。

---

## ✅ 功能特性

- 📁 文件上传与打印（支持图片、PDF、Word、Excel、TXT）
- 🖨️ 多打印机选择
- 🖼️ 图像自动居中、高分辨率打印
- 📄 扫描文件为 PNG 或 PDF（NAPS2 控制）
- 🔔 打印成功提示音
- 📂 上传文件预览 / 删除
- ⬇️ 扫描结果下载 / 预览

---

## ⚙️ 安装指南

### 📌 一、系统要求

- 操作系统：Windows 10（管理员权限运行）
- Python：建议 3.9+（64 位）

### 📦 二、安装依赖库

请先安装以下 Python 库：

```bash
pip install flask pillow pywin32
```

如果你使用 Excel / Word 打印，请确保已安装 Office 2010，其它版本请自行测试， 并使用 32/64 位对应的 Python。

---

### 🧩 三、后端依赖软件

#### 1. 📄 [SumatraPDF](https://www.sumatrapdfreader.org/free-pdf-reader.html)

用于后台静默打印 PDF 文件。

- 安装后默认路径示例：
  ```
  C:\Users\<你的用户名>\AppData\Local\SumatraPDF\SumatraPDF.exe
  ```

#### 2. 📠 [NAPS2 Console](https://www.naps2.com/)

用于命令行扫描图像或 PDF。

- 默认路径：
  ```
  C:\Program Files\NAPS2\NAPS2.Console.exe
  ```

---

## ▶️ 启动方法

在项目根目录运行：

```bash
python app.py
```

浏览器访问：http://localhost:5000

---

## 📁 项目结构

```
printser/
├── app.py                  # 后端主程序
├── templates/
│   ├── index.html          # 首页：上传与打印界面
│   └── preview.html        # 扫描结果预览页面
├── static/
│   └── style.css           # 页面样式（可选）
├── uploads/                # 上传文件目录
└── scans/                  # 扫描结果目录
```

---

## 📤 支持文件类型

| 类型     | 扩展名                          |
|----------|---------------------------------|
| 图像     | .jpg, .jpeg, .png, .bmp         |
| 文档     | .pdf, .doc, .docx, .xls, .xlsx, .txt |

---

## 🛠️ 使用技巧与注意事项

- **推荐用管理员权限运行**，确保有访问打印机、扫描仪权限。
- Word 和 Excel 打印使用 `win32com.client`，启动较慢，避免频繁操作。
- 处理 Excel 和 Word 时，如果失败可能需杀死残留 `winword.exe` 或 `excel.exe`。
- 扫描部分默认使用最近一次在 NAPS2 中配置的设备和设置，如需切换请先在 GUI 设置保存。

---

## 📬 联系与支持

如需帮助或功能扩展，请联系开发者或提交 Issue。
