# Ollama 本地模型使用指南

## 一、安装 Ollama

### Windows

1. **下载安装包**
   - 官网下载：https://ollama.com/download
   - 下载 `OllamaSetup.exe`（约 300MB）

2. **安装**
   - 双击运行，一路默认安装即可
   - 安装完成后 Ollama 会自动在后台运行（系统托盘图标）

3. **验证安装**
   ```bash
   # 打开命令提示符（Win+R → cmd）
   ollama --version
   # 应输出类似: ollama version 0.x.x
   ```

### macOS

```bash
curl -fsSL https://ollama.com/install.sh | sh
```

### Linux

```bash
curl -fsSL https://ollama.com/install.sh | sh
```

---

## 二、拉取模型

Ollama 安装完成后，需要拉取一个模型才能使用。

### 推荐模型

| 模型 | 大小 | 适用场景 | 推荐指数 |
|------|------|----------|----------|
| `qwen2.5:7b` | ~5GB | 通用文本理解，性价比最高 | ⭐⭐⭐⭐⭐ |
| `qwen2.5:3b` | ~2GB | 低配置电脑（4GB 内存即可） | ⭐⭐⭐⭐ |
| `qwen2.5:14b` | ~9GB | 更强理解能力（需 16GB+ 内存） | ⭐⭐⭐ |
| `deepseek-r1:7b` | ~5GB | 思考能力强 | ⭐⭐⭐⭐ |
| `llama3.2:3b` | ~2GB | 极低配置（2GB 内存即可） | ⭐⭐⭐ |

### 拉取命令

```bash
# 通用推荐（约 5GB）
ollama pull qwen2.5:7b

# 低配置电脑（约 2GB）
ollama pull qwen2.5:3b

# 其他模型
ollama pull deepseek-r1:7b
ollama pull llama3.2:3b
```

> **注意**：模型文件较大（2-9GB），下载需要一些时间，请确保网络稳定。

### 查看已拉取的模型

```bash
ollama list
```

---

## 三、启动 Ollama 服务

Ollama 安装后默认开机自启，服务地址为 `http://localhost:11434`。

### 手动启动/停止

```bash
# 启动 Ollama 服务
ollama serve

# 停止（Windows：右键系统托盘图标 → Quit）
# 停止（Linux/Mac：Ctrl+C）
```

### 验证服务是否运行

```bash
# 在浏览器打开，或使用 curl：
curl http://localhost:11434
# 应返回: Ollama is running
```

---

## 四、与软件连接

### 软件端设置

1. 打开软件，切换到 **AI内容命名** 标签页
2. AI 模式选择：**本地模型 (Ollama)**
3. 配置参数：

| 设置项 | 说明 | 默认值 |
|--------|------|--------|
| **模型** | 选择你拉取的模型 | qwen2.5:7b |
| **API 地址** | Ollama 服务地址 | `http://localhost:11434` |
| **温度** | 创意度（0=精确，1=创意） | 0.7 |
| **最大 token** | 生成文件名字数上限 | 100 |

> **默认地址 `http://localhost:11434` 不需要修改**，除非你：
> - 在远程服务器上运行 Ollama（改为服务器 IP）
> - 修改了 Ollama 默认端口

### 测试连接

软件会在调用时自动测试连接。如果出现问题：

```bash
# 1. 确认 Ollama 正在运行
ollama ps

# 2. 确认模型已拉取
ollama list

# 3. 测试 API（在浏览器访问）
http://localhost:11434

# 4. 测试模型生成
curl http://localhost:11434/v1/chat/completions \
  -H "Content-Type: application/json" \
  -d '{
    "model": "qwen2.5:7b",
    "messages": [{"role": "user", "content": "你好"}],
    "stream": false
  }'
```

---

## 五、故障排查

### 常见问题

| 问题 | 原因 | 解决 |
|------|------|------|
| 软件提示"本地模型调用失败" | Ollama 未运行 | 启动 Ollama |
| 软件提示"连接被拒绝" | API 地址不对 | 检查地址是否为 `http://localhost:11434` |
| 生成文件名很慢 | 模型推理慢 | 换小模型（如 `qwen2.5:3b`） |
| 生成的文件名不准确 | 模型理解能力不够 | 换大模型（如 `qwen2.5:14b`） |
| 内存不足 | 模型太大 | 换小模型，或关闭其他程序 |

### 查看 Ollama 日志

```bash
# Windows
# Ollama 日志在: %LOCALAPPDATA%\Ollama\server.log

# macOS
cat ~/.ollama/logs/server.log

# Linux
journalctl -u ollama
```

### 重启 Ollama

```bash
# Windows
# 右键系统托盘 Ollama 图标 → Quit → 重新从开始菜单启动

# Linux/macOS
systemctl restart ollama   # Linux
# 或
ollama serve               # 前台运行
```

---

## 六、性能建议

### 硬件要求

| 模型 | 最低内存 | 推荐内存 | GPU 显存 |
|------|---------|---------|---------|
| qwen2.5:3b | 4GB | 8GB | 2GB |
| qwen2.5:7b | 8GB | 16GB | 4GB |
| qwen2.5:14b | 16GB | 32GB | 8GB |

### 加速建议

1. **使用 GPU 加速**：NVIDIA GPU 可大幅提升速度
2. **换小模型**：CPU 推理时，`qwen2.5:3b` 比 `7b` 快 2-3 倍
3. **减少并发**：AI内容命名中设置为"单文件逐个"模式
4. **限制内容长度**：设为"500 字"减少推理量

### 批量处理建议

- **CPU 用户**：推荐使用"单文件逐个"模式 + "500 字"内容限制
- **GPU 用户**：可以使用"批量合并"模式 + "自动"内容限制
- **大批量文件（100+）**：建议先小批量测试，确认效果后再全量处理

---

## 七、高级：自定义 Ollama 配置

### 修改端口

```bash
# Windows 设置环境变量
setx OLLAMA_HOST "0.0.0.0:11434"

# Linux/Mac
export OLLAMA_HOST="0.0.0.0:11434"

# 重启 Ollama 生效
```

### 允许局域网访问（多电脑共享）

```bash
setx OLLAMA_HOST "0.0.0.0:11434"
```

然后在其他电脑的软件中设置 API 地址为 `http://服务器IP:11434`

### 使用自定义模型

软件支持列表中未列出的模型也可直接使用：
1. 在 Ollama 中拉取任意模型：`ollama pull 你的模型名`
2. 在 AI 内容命名页点击 **「🔍 检测本地模型」**，软件会自动发现并列出 Ollama 中已安装的全部模型
3. 从下拉列表中选择该模型即可使用，无需修改任何代码

---

## 八、卸载 Ollama

```bash
# Windows
# 设置 → 应用 → Ollama → 卸载

# macOS
brew uninstall ollama

# Linux
sudo apt remove ollama
```
