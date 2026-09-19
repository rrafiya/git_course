# 雷霆战机 · Thunder Fighter

一个用原生 HTML5 Canvas + JavaScript 编写的竖版飞行射击小游戏。**单文件、零依赖、开箱即玩。**

## 在线试玩

如果已在仓库 Settings → Pages 中开启 GitHub Pages：

👉 **https://rrafiya.github.io/git_course/**

## 本地运行

直接双击 `index.html` 用浏览器打开即可，无需安装任何东西、无需构建。

如果想用本地服务器（可选）：

```bash
python -m http.server 8000
# 然后浏览器访问 http://localhost:8000
```

## 操作方式

| 操作 | 按键 |
|---|---|
| 移动 | `WASD` / 方向键 / **按住鼠标拖动** |
| 射击 | 按住 `空格` 或 **鼠标左键**（自动连射，无需连点） |
| 暂停 | `P` |
| 重新开始 | `R` |
| 静音开关 | `M` |

> **关于鼠标/触屏操作**：采用「相对拖动」——按住后飞船跟随指针的**位移量**移动，松开即停止。
> 指针单纯悬停或划过屏幕不会带动飞船，也不存在自动漂移。
> 手机 / 平板上按住屏幕拖动即可移动并自动开火。

## 玩法说明

- **火力升级**：初始单发，最高 3 级。3 级为双主炮 + 双斜射。
- **道具掉落**：

  | 道具 | 效果 |
  |---|---|
  | `W` 黄色 | 火力等级 +1（满级则转化为 300 分） |
  | `S` 蓝色 | 护盾 +1，可抵挡一次伤害 |
  | `R` 紫色 | 速射，大幅缩短射击间隔 |
  | `+` 绿色 | 回复 1 点生命 |

- **生命**：3 点。受击后有短暂无敌闪烁，此时不会连续掉血。
- **波次**：敌人强度和数量随波次递增；**每 5 波出现 BOSS**，BOSS 有四套弹幕模式（扇形、瞄准三连、双炮口、环形）。
- **计分**：击落敌机得分，种类不同分值不同（小机 100 / 精英 180 / 重甲 320 / BOSS 2500）。最高分通过 `localStorage` 保存在本地。

## 技术实现

- 纯原生 JavaScript，无任何框架或第三方库
- 逻辑分辨率固定 540×960，通过 `ResizeObserver` 测量容器并等比缩放，保证**任何窗口比例下画面都完整可见、不被裁切**
- `requestAnimationFrame` 主循环，基于真实帧间隔计算 `dt`，不同刷新率下速度一致
- Canvas 2D 程序化绘制：所有飞船、粒子、星云均为代码绘制，无图片资源
- 星空视差滚动、粒子爆炸、冲击波、屏幕震动、受击顿帧、光晕叠加
- Web Audio API 实时合成音效，无需音频文件
- 相对拖动输入：仅响应指针位移，无输入时飞船绝对静止；失焦/切后台自动清空输入状态

## 文件说明

```
index.html      # 游戏本体（结构 + 样式 + 逻辑，全部在这一个文件里）
diagnose.html   # 显示问题排查工具（遇到画面异常时打开它）
README.md       # 本说明
.gitignore      # 忽略规则
```

## 布局规则（改动前请先读）

画布尺寸由 `resize()` 统一计算，几条硬性约束不要破坏：

1. **画布尺寸绝不超过可见区域** —— 缩放结果会再做一次硬夹紧（`host.w` / `host.h` / `maxW` / `maxH`）。
2. **不得出现"提前 return 而不缩小画布"的分支** —— `<canvas>` 的 HTML 初始尺寸是
   540×960，一旦某次 resize 直接被 return，画布就会停留在全尺寸并溢出小窗口。
   这是历史上真实出现过的缺陷。
3. **垂直位置要避开底部提示条** —— 底部 `#hint`（约 18px 高 + 6px 边距）占据窗口
   底部，`resize()` 会按它的实测高度预留空间，并把画布居中到剩余区域内。
   CSS 用 `translateX(-50%)` 只做水平居中，`top` 由 JS 设置；若改回
   `translate(-50%,-50%)` 会导致垂直定位偏差半个画布高度。
4. **每 700ms 的自愈检查**（`selfCheck`）会在检测到画布越界时自动重新计算，
   这是最后一道保险。

## 画面异常排查

如果遇到「内容显示不全」「自机或敌人跑到窗口外看不到」这类问题，
用浏览器打开 **`diagnose.html`**，它会：

- 用 iframe 加载游戏并实时测量画布的实际显示范围
- 用红色描边框出画布边界，并逐项检查是否超出可见区域
- 采样画布像素，算出实际绘制内容的包围盒
- 给出结论并提供「复制诊断报告」按钮

把复制出来的报告贴出来即可定位问题。

## 在本机用真实 Chrome 验证渲染

本项目不改用任何测试框架，但可以用无头 Chrome + CDP 做真实的端到端验证
（真实布局测量 + 截图）。**注意本机环境限制**：

> 沙箱默认禁止创建**命名管道**，而 Chrome 的 Mojo IPC 依赖它，直接启动会失败：
> `FATAL:mojo/public/cpp/platform/platform_channel.cc: Check failed: 拒绝访问 (0x5)`。
> 解决方式是放宽该命令的沙箱权限（full access）后再启动 Chrome。

另外 **不要用 Node 的 `child_process.spawn` 去启动浏览器**（管道 stdio 在该环境下
会抛 `EPERM`）；应改用 PowerShell 的 `Start-Process` 启动，再用 Node 通过
WebSocket 连 CDP。

```powershell
# 1) 启动无头 Chrome（需要放宽沙箱权限）
Start-Process "C:\Program Files\Google\Chrome\Application\chrome.exe" -ArgumentList @(
  "--headless=new","--remote-debugging-port=9501","--user-data-dir=<临时目录>",
  "--no-first-run","--no-default-browser-check","--no-sandbox","--disable-gpu","about:blank"
)
# 2) 用 Node 连接 http://127.0.0.1:9501/json/new?<file-url> 并走 CDP 驱动
```

## 许可

仅用于学习交流，随意取用。
