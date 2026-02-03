# 项目开发规范

## Shell 命令重定向规范

**重要**：本项目在跨平台环境（Windows + Git Bash）下运行，使用 Bash 工具时必须遵循以下规范：

### ✅ 正确用法

```bash
# 静默错误输出（正确，适用于所有 Unix-like shell）
command 2>/dev/null

# 静默标准输出
command >/dev/null

# 静默所有输出
command &>/dev/null
```

### ❌ 错误用法

```bash
# 错误！Windows 风格，在 Git Bash 中会创建名为 "nul" 的文件
command 2>nul

# 错误！同样会创建 "nul" 文件
command >nul
```

### 原因说明

- `/dev/null` 是 Unix-like 系统的标准空设备
- `nul` 或 `NUL` 是 Windows 的空设备，但在 Git Bash/MSYS2 环境中会被当作普通文件名
- 混用会导致创建不需要的文件（如 `nul` 文件）

## 其他注意事项

- Windows 批处理文件（.bat）中使用 `>nul` 是正确的
- 但在所有通过 Bash 工具执行的命令中，必须使用 `>/dev/null`
