---
name: strategy-slide-architect
description: Generate a 10-30 slide markdown or Marp deck when the user asks to convert a business strategy document into an executive narrative, including slide-by-slide key messages, evidence structure, and clear mapping to Core/Why/What/How.
---

# Goal
戦略ドキュメントを、経営陣・投資家・事業責任者に通用するスライド構成へ変換する。  
出力は「意思決定に必要な論点」と「伝達順序」を明確にし、Core/Why/What/Howを一貫接続する。

# Steps
1. 入力資料から、Core/Why/What/Howごとに主張と根拠を抽出する。
2. 想定オーディエンス（経営会議、取締役会、投資家など）を明示する。
3. 冒頭でWhy Nowを提示し、中盤でWhat、終盤でHowと意思決定事項を配置する。
4. 各スライドに以下を記載する。
   - Slide title
   - Key message（1文）
   - Evidence（定量・事例・比較）
   - Mapping（Core/Why/What/How）
5. スライドのトーンを「断定→根拠→次アクション」の順で統一する。
6. テンプレートが必要なときは `slide_template.md` を読み、形式を揃える。

# Examples
## Input
- `v1_nexpro_strategy.md` のような戦略本文
- KPI表、競合比較、ロードマップ

## Output
- `*_slide_outline.md`（10〜15枚または指定枚数）
- 必要に応じてMarp形式本文（`---` 区切り）

## Quality bar
- 1スライド1メッセージであること。
- 全スライドがCore/Why/What/Howのどこかに紐づくこと。
- 最終スライドに意思決定事項と90日アクションを入れること。
