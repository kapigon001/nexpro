---
name: strategy-consistency-reviewer
description: Review strategy documents when the user asks for logic checks, contradiction detection, or investment-readiness validation across Core/Why/What/How, and return concrete fixes for gaps in assumptions, KPIs, and execution links.
---

# Goal
戦略文書の階層整合性を監査し、Core/Why/What/Howの飛躍を修正可能な形で提示する。  
レビュー結果は、意思決定者が「この戦略は実行できる」と判断できる水準まで具体化する。

# Steps
1. Coreを抽出し、Vision・Mission・事業戦略が同じ方向を向いているか確認する。
2. Whyを確認し、対象顧客・課題・市場変化・競争環境がCoreを正当化しているか判定する。
3. Whatを確認し、提供価値・収益モデル・KPIがWhyの課題解決に直結しているか検証する。
4. Howを確認し、UI/実装/GTM/体制がWhatを期限内に実現できる設計か評価する。
5. 不整合を以下の形式で出力する。
   - Gap（何が欠けているか）
   - Impact（どの意思決定に影響するか）
   - Fix（どの文・指標・施策をどう修正するか）
6. 最後に「Go / Conditional Go / No Go」の判定と条件を提示する。

# Examples
## Input
- 戦略本文（例: `v1_nexpro_strategy.md`）
- スライドアウトライン（例: `v1_nexpro_slide_outline.md`）

## Output
- 整合性レビュー表（Core→Why→What→How）
- 優先修正TOP5
- 実行前チェックリスト（指標・体制・マイルストーン）

## Quality bar
- 各指摘に必ずImpactとFixをつけること。
- KPIの欠落は具体指標（定義・計測頻度）まで補うこと。
- GTMとプロダクトが分断している場合は接続施策を明示すること。
