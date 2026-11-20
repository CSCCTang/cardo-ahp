// cardo-ahp 的 Host 半：把包内 skill 注册到 ctx.skills。
// 依据《dsh 插件开发标准（现行版）》5.1、6.1。
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const name = 'cardo-ahp'
export const inject = ['skills']

const skillDir = join(dirname(fileURLToPath(import.meta.url)), 'skills', 'cardo-ahp')

function skillBody(raw) {
  // 去掉 YAML frontmatter：模型侧只接正文
  return raw.replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n/, '')
}

export function apply(ctx) {
  ctx.skills.register({
    name: 'cardo-ahp',
    description: "Cardo AHP 投资决策量化流水线：专家问卷、权重与一致性、方案评分与成果装配。",
    whenToUse: "用户要做 AHP 专家问卷、算权重与一致性、做方案评分并产出《AHP权重表》《AHP方案评分结论》《AHP权重与评分结论》《AHP依据对照表》时使用。",
    content: skillBody(readFileSync(join(skillDir, 'SKILL.md'), 'utf8')),
    resourceBase: { kind: 'directory', path: skillDir },
  })
}
