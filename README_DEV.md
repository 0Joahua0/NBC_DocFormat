## 关于 custom_settings.json

仓库里这个文件是**源码运行模式（`python NBC_DocFormat.py` / `bash install.sh`）的默认配置**，
内容与代码里的 DEFAULT_CUSTOM_SETTINGS 保持一致。

### 给开发者的提示

该文件会随源码一起提交，作为源码运行模式的默认配置。**如果你在调试时通过 GUI 保存了自定义设置，
会改动本地这个文件；提交前请确认它仍然是干净的默认值**，避免把个人调试配置误提交。

如果你确实要修改这份"默认配置"（比如调整 DEFAULT_CUSTOM_SETTINGS 后想同步到这里），
请一并更新代码中的默认值和相关测试。

### 用户配置文件的实际位置

- Windows/Linux 打包发布版：exe 同目录
- macOS 打包发布版：~/Library/Application Support/NBC_DocFormat/
- 开发模式（python 直接运行）：项目根目录（即这个文件本身）
