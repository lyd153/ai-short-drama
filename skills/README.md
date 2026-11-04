# 技能库

可安装、可复用的创作技能。每个技能是一个目录，内含 `SKILL.md`（入口）与 `references/`（手册）、`templates/`（模板）。

## 技能清单

| 技能 | 用途 | 触发场景 |
| --- | --- | --- |
| [novel-to-screenplay](novel-to-screenplay/SKILL.md) | 小说 → 剧本 → 15 秒单元 → 分镜 → 逐镜提示词 → 负面词 → 质检 | 要做改编、拆场景、写分镜、生成提示词、选运镜、补负面词 |

## 安装到 DSH

技能目录复制到用户技能目录即可被识别：

```powershell
# Windows
Copy-Item -Recurse -Force .\skills\novel-to-screenplay "$env:USERPROFILE\.dsh\skills\novel-to-screenplay"
```

```bash
# macOS / Linux
cp -R skills/novel-to-screenplay ~/.dsh/skills/
```

装好后在对话里说「用 novel-to-screenplay 把这章改成 15 秒单元分镜」即可触发。

## 技能写作规范

1. `SKILL.md` 顶部必须有 YAML frontmatter：`name`、`description`。
2. `description` 写清**什么时候用**（触发条件），不是写这是什么。
3. 正文先给流程（可执行步骤），再给参考文件索引；细节放 `references/`，不要全塞在入口。
4. 每个技能自带「质检清单」与「失败处理」，否则不可交付。
5. 引用的外部项目与教程，写清来源，不复制原文。

## 新增技能

```text
skills/
└─ 技能名/
   ├─ SKILL.md
   ├─ references/
   └─ templates/
```

命名用 kebab-case（如 `novel-to-screenplay`），中文名放进 `SKILL.md` 标题。
