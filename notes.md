# Notes — DakeSCI behavior guesses

## What the website confirms
- ¥99/device PPT add-in, 兑换码→机器码→激活码 licensing
- Branding: "AI绘图网站" but actual shipped features are PPT shape ops

## Behavior interpretation per feature
| # | Feature | My interpretation |
|---|---------|-------------------|
| 1 | PNG 空白边裁除 | Detect non-white bbox, crop, optionally batch a folder |
| 2 | 生成表格 | CSV/Markdown/list → native PPT table on a slide |
| 3 | 层状结构 | Stack labeled rectangles vertically with gap+arrow (architecture layer cake) |
| 4 | 矩阵逐行偏移 | RxC grid of shapes where every other row shifts by Δx (brick pattern) |
| 5 | 表格排版保持图片比例 | Fill table cell with image preserving aspect (letterbox/pillarbox, no stretch) |
| 6 | 图标补充 | Add bundled icons by name to a slide |
| 7 | 统一安装日期 | Single source of truth for build/release date shown in --version |

## Unknowns deferred
- Exact ribbon UI layout — out of scope (no PPT)
- AI image generation hinted by site name — NOT in 1.2 release notes, skipping
- Licensing flow — NOT in scope of "脚本工具" reproduction
