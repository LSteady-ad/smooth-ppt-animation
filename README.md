# smooth-ppt-animation

## 项目描述（用于搜索 / 发现）

为已有 PowerPoint（`.pptx`）制作高级感动画的通用 AI 工具包：结合 Morph 平滑切换、`!!` 同名元素配对、跨页屏外副本飞入 / 飞出、页内入场动画、z-order 层级修正与自动校验。适用于“给 PPT 加动画 / 平滑切换 / Morph / 炫酷切换 / 元素飞入飞出 / 每页精美动画 / 基于最新版重新上动画”等需求；支持大面积内容随平滑移动、细节文字分批进入、每页 4 秒内完成、图片跨页关系修复。

该工具包沉淀了多类演示文稿动画制作中的常见失败模式，把容易出错的配对、层级、图片引用和时序问题都做成了可复用的脚本与检查项。可被各类具备文件读写与脚本执行能力的 AI 助手直接调用；若工具支持 `SKILL.md` 约定，也可作为技能目录加载。

## 它能做什么

- 把大标题、色块、图片、卡片等主要内容做成跨页平滑移动
- 把下一页元素复制到上一页屏外，实现“平滑飞入”
- 把当前页主要元素复制到下一页屏外，实现“平滑飞出”
- 对说明文字、评级字等细节使用淡入 / 缩放 / 上浮，节奏可统一、顺序、波浪或接力
- 自动修复跨页图片 `r:embed` 引用，防止大图在平滑中变成错误形状
- 自动对齐 z-order，避免重叠元素在切换中间帧闪烁
- 幂等重做：用户改过文字字体后，可基于最新版重新上动画，不覆盖原文件
- 自动校验：文字一致性、图片引用、动画数量、首动画延迟、总时长、层级冲突

## 安装方式

无需构建步骤。任选一种方式：

1. 将本项目目录放入你的 AI 工具所支持的技能 / 工具目录；支持 `SKILL.md` 约定的工具会自动识别。
2. 直接下载或克隆本仓库，在任意项目中调用 `scripts/` 下的脚本。
3. 若使用发布包，解压 `smooth-ppt-animation.zip` 后按上面两种方式放置即可。

## 使用方式

安装后，可以这样对 AI 助手说：

> 给这个 PPT 加上高级感的平滑动画，每个元素都要有动画
>
> 基于最新版重新上动画，我改过字体文字
>
> 用 Morph 做跨页飞入飞出，细节文字再逐个进入，每页控制在 4 秒内

AI 助手会按完整流程执行：检查 → 复制只读基准 → 构建 → 校验 → 交付新版本。

## 手动命令行用法

```bash
# 1. 检查原始演示文稿结构
python scripts/inspect_deck.py input.pptx --page 4

# 2. 构建动画（源文件不会被修改）
python scripts/build_morph_anim.py \
  --src input.pptx \
  --out input_animated.pptx \
  --ref clean_named_deck.pptx \   # 可选：已有 !! 命名体系的参考稿
  --tpl deck_with_templates.pptx  # 可选：含 presetID 10/42/55/2/53 的模板稿

# 3. 校验结果
python scripts/verify.py input_animated.pptx --src input.pptx
```

## 目录结构

```text
smooth-ppt-animation/
├── SKILL.md                     入口说明与完整工作流
├── README.md                    本文件
├── scripts/
│   ├── build_morph_anim.py       生产构建脚本
│   ├── verify.py                 完整性 / 一致性校验脚本
│   └── inspect_deck.py           构建前结构检查脚本
└── references/
    ├── concepts.md               Morph 原理、命名、动画类型与时序
    ├── pitfalls.md               常见失败案例与修复方法
    └── config.md                 参数配置参考
```

## 已编码的关键规则

- 跨页配对只可靠依赖 `!!name`；没有 `!!` 的元素会被 PowerPoint 智能配错
- `r:embed` 是页本地关系，复制图片到别的页必须重新映射
- 第一个动画延迟必须为 0，否则平滑结束后会“等一下才出字”
- 同名配对元素不仅形状要合适，z-order 百分位也要接近
- 大面积色块即便外形不完全一样，也可以参与跨页平滑；小装饰元素则页唯一化，避免误配
- 绝不覆盖用户可能编辑过的文件，始终输出新版本
- 源文件和输出文件必须是两个不同的绝对路径

完整失败模式与修复细节见 `references/pitfalls.md`。

## 说明

本工具包来自实战经验沉淀。`!!` 命名是 PowerPoint Morph 的官方特性，可参考：

https://support.microsoft.com/en-us/office/morph-transition-tips-and-tricks
