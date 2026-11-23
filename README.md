# AI PPT 数字人教学助手

在 **不改动原 PPT 任何内容、版式、原有动画** 的前提下，为指定幻灯片嵌入统一的数字人动态形象（说话动画 + 结束自动切回静态待机闭嘴图，彻底无重影）与高清配音，支持一键导出新 `.pptx` 与课堂智能放映（翻页/回退自动复位重播）。

---

## 一、给客户的极简下载与安装步骤（Windows 10 / 11）

> **运行前提**：电脑已安装 **Microsoft PowerPoint**（桌面版，非仅 WPS）。

### 方式 1：直接下载绿色免安装版（最简单，推荐）

1. **下载压缩包**：点击进入最新版下载页，下载 `AI-PPT-Digital-Human-Win64-Portable.zip`（已内置独立 Python 与 FFmpeg，无需手动配环境）：  
   👉 **[点此下载最新版便携包（GitHub Releases）](https://github.com/bbwx91697-hash/AI-PPT-Digital-Human/releases/latest/download/AI-PPT-Digital-Human-Win64-Portable.zip)**  
   （打不开直链时可点此进入发布页：[Releases 页面](https://github.com/bbwx91697-hash/AI-PPT-Digital-Human/releases/latest)）
2. **解压到任意磁盘目录**：右键压缩包 → 解压到当前文件夹（建议放 D 盘，如 `D:\AI-PPT-Digital-Human`）。
3. **安装默认高音质语音模型（推荐）**：
   - 打开解压后的文件夹，双击 **`一键安装本地Qwen3高音质模型.bat`**。
   - **有 NVIDIA 独立显卡（RTX 系列等）**：自动走国内镜像安装 CUDA 加速版，10 秒语音约 15~20 秒生成。
   - **没有 NVIDIA 显卡（核显 / AMD 显卡 / 轻薄本）**：自动改为安装轻量化 CPU 版（约 180 MB，无需 2.7 GB 显卡驱动），同样可以运行 Qwen3 高音质模型，生成速度约为音频时长的 **3.4 倍**（10 秒语音约 35 秒，也可随时切换免费微软神经语音）。
   - 全程走阿里云 / 清华 PyPI / ModelScope 国内镜像，无需翻墙，安装完成后会自动清理临时安装包缓存。
4. **启动软件**：双击 **`启动数字人PPT助手.bat`** 即可使用！

### 方式 2：从源码仓库安装

如果你直接下载了 GitHub 源码压缩包（`Code -> Download ZIP`）或使用 `git clone`：
1. 解压后双击 **`一键安装并启动.bat`**（自动创建 `.venv` 并通过清华镜像安装基础依赖）。
2. 再双击 **`一键安装本地Qwen3高音质模型.bat`** 部署本地高音质语音模型。

---

## 二、让「豆包」或 AI 助手帮你下载与安装

如果客户不熟悉 GitHub 下载流程，可以把本仓库链接直接发给 **豆包**，并把下面这段话直接发给它执行或引导：

> **发送给豆包的参考指令**：  
> 请帮我从 GitHub 下载并指导我安装这个软件：`https://github.com/bbwx91697-hash/AI-PPT-Digital-Human`。  
> 请优先让我下载 `Releases` 里的便携压缩包 `AI-PPT-Digital-Human-Win64-Portable.zip`（直链：`https://github.com/bbwx91697-hash/AI-PPT-Digital-Human/releases/latest/download/AI-PPT-Digital-Human-Win64-Portable.zip`），解压到 D 盘后先运行 `一键安装本地Qwen3高音质模型.bat`，再运行 `启动数字人PPT助手.bat`。

### 一键 PowerShell 下载并解压命令（可复制到 Windows 终端运行）

```powershell
$dir = "D:\AI-PPT-Digital-Human"
New-Item -ItemType Directory -Force -Path $dir | Out-Null
Invoke-WebRequest -Uri "https://github.com/bbwx91697-hash/AI-PPT-Digital-Human/releases/latest/download/AI-PPT-Digital-Human-Win64-Portable.zip" -OutFile "$dir\pkg.zip"
Expand-Archive -Path "$dir\pkg.zip" -DestinationPath "D:\" -Force
Remove-Item "$dir\pkg.zip" -Force
explorer $dir
```

---

## 三、常见语音引擎对比

| 语音引擎 | 音质与自然度 | 是否需要显卡/网络/额度 | 适用场景 |
| --- | --- | --- | --- |
| **Qwen3-TTS 0.6B（默认推荐）** | 极高，最接近真人讲课 | 装完可**完全断网离线使用**，永久免额度；有 N 卡最快，无 N 卡也能在 CPU 运行 | 正式公开课、比赛说课、毕业答辩 |
| **微软神经语音（Edge TTS）** | 良好、自然 | 需联网、免费免 Key、CPU 电脑秒出 | 临时赶制、无空间下载本地模型 |
| **Windows 离线语音** | 基础朗读音 | 系统自带、断网即用 | 极简应急预览 |
| **本地音频 / 视频音轨导入** | 原声 | 免费、支持 WAV/MP3/M4A/MP4 等 | 直接导入真人录音 |
| **Gemini 3.8 云端语音** | 极高 | 需自备 Google API Key（Windows DPAPI 本机加密保存） | 已有官方 API 额度的用户 |

---

## 四、使用流程简览

1. 点击右上角 **打开 / 导入 PPTX**（软件会自动复制一份进 `projects/`，**绝不修改原始课件**）。
2. 在左侧勾选需要开口讲解的幻灯片，在中间预览区直接用**鼠标拖动 / 四角缩放**调整数字人大小与位置（支持一键应用到所有已启用页面）。
3. 在右侧填写台词并点击 **一键生成本页语音**（或批量生成，自动按页缓存，改某页不重跑其他页）。
4. 选择 **自动播放** 或 **点击人物播放**，点击 **一键生成并导出新 PPTX**。
5. 导出后可直接点击 **启动智能放映**，翻页或回退均会自动重播当前页数字人。

---

## 五、核心文件说明

| 文件/目录 | 用途 |
| --- | --- |
| `启动数字人PPT助手.bat` | 一键启动桌面软件（优先使用便携包内置运行时） |
| `一键安装并启动.bat` | 纯源码下载时的基础环境一键安装与启动脚本 |
| `一键安装本地Qwen3高音质模型.bat` | 一键安装 Qwen3-TTS 0.6B（自动识别 N 卡 CUDA / 无 N 卡 CPU 版，国内镜像） |
| `src/app.py` | Windows 高 DPI 桌面主界面 |
| `src/qwen3_worker.py` / `src/tts_service.py` | 多引擎语音服务与防卡死/首字缓冲常驻推理守护 |
| `src/openxml_merge.py` / `src/timing_builder.py` | Open XML 原动画保真合并与动/静图严格互斥时序 |
| `src/embed_media.ps1` / `src/smart_slideshow.ps1` | PowerPoint COM 媒体注入与智能放映控制器 |
| `assets/avatar/` | 默认透明背景数字人 GIF 与第 0 帧静态待机 PNG |
