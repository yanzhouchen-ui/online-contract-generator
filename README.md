# 线上合同制作器

基于 Streamlit 的 Web 版合同批量生成工具。

## 功能

- 上传 Excel 数据文件
- 使用内置模板或上传自定义模板
- 批量生成合同文档
- 打包下载为 ZIP

## 本地运行

```bash
# 安装依赖
pip install -r requirements.txt

# 启动应用
streamlit run app.py
```

## 部署到 Streamlit Cloud

1. Fork 或推送此仓库到 GitHub
2. 访问 [share.streamlit.io](https://share.streamlit.io)
3. 连接 GitHub 仓库
4. 点击 Deploy

## 使用说明

1. 准备 Excel 数据文件（包含合同字段）
2. 打开应用，上传 Excel 文件
3. 选择模板（内置或自定义）
4. 设置签署日期和价格
5. 点击生成，下载 ZIP 文件

## 模板占位符

- Excel 列名 → `[Column Name]` / `[column name]`
- 日期 → `[Fecha]` / `[fecha]` / `[Date]`
- 价格 → `[precio]` / `[Precio]`
