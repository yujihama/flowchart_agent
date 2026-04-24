# JSOX フローチャート生成エージェント — アーキテクチャ概要

## 1. 目的

J-SOX（金融商品取引法に基づく内部統制報告制度）における **3点セット**
（業務記述書・業務フロー図・RCM）のうち、**業務フロー図** を
自然言語の業務ヒアリング結果や既存資料から半自動で生成するエージェント。

本ドキュメントのスコープは「フローチャート部分」のみ。RCM/業務記述書との
連携はインタフェースとして定義するが、実装は別コンポーネントに委ねる。

### 1.1 J-SOX フローチャートに求められる要件

一般的な業務フロー図と比べ、以下が追加で求められる:

| 要件 | 内容 |
|---|---|
| スイムレーン | 部門・担当者ごとにレーンを分け、責任の所在を明確化 |
| リスク点(R) | 誤謬・不正が生じうる処理点を明示（例: R1, R2） |
| 統制点(C) | リスクに対応する統制活動を明示（例: C1, C2） |
| キーコントロール | 特に重要な統制を識別（◎印など） |
| 証憑・帳票 | どの書類がどの段階で発生/参照されるかを表現 |
| IT 自動化 | 手作業 / IT 自動化 / IT 依存手作業 の区別 |
| 分岐の条件 | 承認/否認、与信 OK/NG など業務判断の条件 |

ユーザ提示の最小サンプルは上表の「スイムレーン」「分岐の条件」をカバー。
他の要素は **データモデル拡張ポイント** として設計に織り込む。

## 2. レイヤー構成

```
┌────────────────────────────────────────────────────────────┐
│  ① 入力層  (Intake)                                         │
│    - 自然言語ヒアリング / 既存Excel / 業務記述書テキスト    │
└───────────────┬────────────────────────────────────────────┘
                │
┌───────────────▼────────────────────────────────────────────┐
│  ② エージェント層  (Agent / LLM)                            │
│    - 業務抽出 → レーン・ノード・エッジ推定                  │
│    - リスク/統制点の発見・提案                              │
│    - ユーザとの対話的refinement                             │
└───────────────┬────────────────────────────────────────────┘
                │            生成 / 編集
┌───────────────▼────────────────────────────────────────────┐
│  ③ ドメインモデル層  (Flow DSL)                             │
│    - Flow / Lane / Node / Edge (+ Risk / Control / Doc)     │
│    - バリデーション (到達可能性・孤立ノード・循環等)        │
└───────────────┬────────────────────────────────────────────┘
                │            シリアライズ
┌───────────────▼────────────────────────────────────────────┐
│  ④ レンダリング層  (Renderer)                               │
│    - Mermaid / Graphviz(DOT) / draw.io(XML) / SVG / PNG    │
│    - レイアウト: レーン別 Y 座標割り当て, トポロジカル配置 │
└───────────────┬────────────────────────────────────────────┘
                │
┌───────────────▼────────────────────────────────────────────┐
│  ⑤ 永続化 / 連携層  (Persistence & Export)                  │
│    - JSON/YAML でのフロー保存（diff可能）                   │
│    - RCM, 業務記述書との ID リンク                          │
└────────────────────────────────────────────────────────────┘
```

## 3. ドメインモデル

### 3.1 最小モデル（ユーザ提示）

```python
Flow(lanes=[Lane...], nodes=[Node...], edges=[Edge...])

Lane(id, name)
Node(id, lane_id, label, type)        # type: start | end | decision | task(default)
Edge(from_id, to_id, condition=None)
```

### 3.2 J-SOX 拡張（段階的に追加）

```python
# ノード種別を拡張
Node.type ∈ {start, end, task, decision, subprocess, document, manual, it_auto, it_dependent}

# リスク・統制
Risk(id, node_id, description, level)              # 例: R1, R2
Control(id, node_id, description, key=False,       # ◎キーコントロール
        nature={preventive|detective},
        method={manual|it_auto|it_dependent},
        frequency)
RiskControlLink(risk_id, control_id)

# 証憑・帳票
Document(id, name, format)                          # 注文書, 検収書 など
NodeDocument(node_id, document_id, action={create|read|update})
```

設計ポイント:
- **ID は人間可読**(`R1`, `C1`, `DOC-ORDER`) にして RCM と相互参照可能に
- 拡張はすべて **オプショナル**。最小モデルだけで描画は成立
- モデルは **Pydantic** / **dataclass** で宣言し、JSON/YAML への往復を可能にする

## 4. エージェントのワークフロー

```
  ユーザ                            エージェント
    │                                   │
    │  「受注から売上計上までの流れ」  │
    ├──────────────────────────────────▶│
    │                                   │  ① 業務分解 (LLM)
    │                                   │    - 登場部門 → Lane
    │                                   │    - 工程   → Node
    │                                   │    - 順序   → Edge
    │                                   │
    │ ◀─── draft Flow (DSL)  ───────────│
    │                                   │
    │  「与信チェックを追加して」       │
    ├──────────────────────────────────▶│  ② 差分編集 (tool-use)
    │                                   │    add_node / add_edge
    │ ◀─── updated Flow ────────────────│
    │                                   │
    │  「リスクと統制を洗い出して」     │
    ├──────────────────────────────────▶│  ③ R/C 提案 (LLM)
    │                                   │    Node 単位で候補を生成
    │ ◀─── Risk/Control 候補リスト ─────│
    │                                   │
    │  「OK、Mermaid で出して」         │
    ├──────────────────────────────────▶│  ④ Render
    │ ◀─── mermaid / svg ───────────────│
```

### 4.1 エージェントが備えるべきツール（関数呼び出し）

| ツール | 役割 |
|---|---|
| `add_lane / add_node / add_edge` | 要素追加 |
| `update_node / move_node` | 属性変更・レーン移動 |
| `remove_*` | 要素削除（参照整合性チェック付） |
| `attach_risk / attach_control` | R/C の付与 |
| `validate(flow)` | 静的チェック（start/end 存在、孤立、不達など） |
| `render(flow, format)` | Mermaid/DOT/drawio 等へ変換 |
| `suggest_risks(node)` | J-SOX観点でのリスク候補生成（LLM） |

これらは **ドメインモデルを変更する唯一の経路** とし、エージェントが
任意の JSON を吐いてそのまま保存する設計は避ける（不正状態を作りにくくする）。

## 5. バリデーション

レンダリング前に必ず実行:

1. `start` ノードが 1 つ以上、`end` ノードが 1 つ以上
2. すべての `Node.lane_id` が既存の Lane に存在
3. すべての `Edge.from_id / to_id` が既存の Node に存在
4. `decision` ノードから出るエッジは 2 本以上あり、各 `condition` は非空
5. 孤立ノード（どの edge にも現れない）の検出 → 警告
6. `start` から全ノードへ到達可能か / 全ノードから `end` へ到達可能か
7. キーコントロールが 1 つ以上存在するか（J-SOX 推奨・警告レベル）

## 6. レンダリング

### 6.1 Mermaid（第一選択）

- 長所: テキストで diff しやすい、GitHub/Notion 等で直接表示
- 課題: スイムレーンは `subgraph` で近似（厳密なレーン表現ではない）

### 6.2 Graphviz (DOT) + クラスタ

- レーンを `cluster_*` で表現。自動レイアウトが強い。

### 6.3 draw.io (XML) / SVG 直接生成

- 監査提出用の清書。**レーン幅・座標を厳密に制御** したい場合はこちら。
- Y 座標 = レーンのインデックス、X 座標 = トポロジカル順序 で配置。

どのレンダラを使うかは `render(flow, format="mermaid"|"dot"|"drawio"|"svg")`
で切替え。**ドメインモデル → 中間 IR → 各フォーマット** の流れにして、
レンダラ追加時にモデル側を触らなくて済むようにする。

## 7. 提示サンプルの解釈

```python
flow = Flow(
    lanes=[Lane("sales","営業部"), Lane("wh","倉庫"), Lane("acc","経理部")],
    nodes=[
        Node("n1","sales","受注",     type="start"),
        Node("n2","sales","注文入力"),
        Node("n3","sales","与信OK?",  type="decision"),
        Node("n4","wh",   "出荷"),
        Node("n5","acc",  "売上計上"),
        Node("n6","acc",  "完了",     type="end"),
    ],
    edges=[
        Edge("n1","n2"),
        Edge("n2","n3"),
        Edge("n3","n4", condition="OK"),
        Edge("n3","n1", condition="NG"),   # 差戻しループ
        Edge("n4","n5"),
        Edge("n5","n6"),
    ],
)
```

このサンプルは **最小の骨格**。次ステップで以下を重ねる想定:

- `n3`（与信 OK?）に `R1: 与信限度超過の見落とし`、`C1: 与信マスタとの自動突合（◎）`
- `n4`（出荷）に `DOC: 出荷指図書 / 納品書`、`C2: 出荷実績と受注の三点照合`
- `n5`（売上計上）に `R2: 期ズレ計上`、`C3: 出荷実績ベースでの自動仕訳（◎）`

## 8. ディレクトリ構成（想定）

```
flowchart_demo/
├── ARCHITECTURE.md              ← 本書
├── pyproject.toml
├── src/jsox_flow/
│   ├── model.py                 ← Flow / Lane / Node / Edge / Risk / Control
│   ├── validate.py              ← バリデーション
│   ├── render/
│   │   ├── mermaid.py
│   │   ├── dot.py
│   │   └── drawio.py
│   ├── agent/
│   │   ├── tools.py             ← add_node 等の tool 定義
│   │   ├── prompts.py           ← システムプロンプト・J-SOX 観点
│   │   └── runner.py            ← Claude Agent SDK 連携
│   └── io/
│       ├── json_io.py
│       └── yaml_io.py
├── examples/
│   └── sales_to_billing.py      ← 提示サンプル
└── tests/
```

## 9. 技術選定メモ

| 項目 | 候補 | 推奨 |
|---|---|---|
| 言語 | Python 3.11+ | Python（Pydantic/LLM SDK が揃う） |
| モデル定義 | Pydantic v2 / dataclass | Pydantic v2（JSON 往復・validation 内蔵） |
| エージェント | Claude Agent SDK / LangGraph | Claude Agent SDK（tool-use が直截） |
| 描画 | Mermaid / Graphviz / drawio | まず Mermaid、監査用に drawio を追加 |
| テスト | pytest + ゴールデンファイル | 各レンダラはスナップショットで固定 |

## 10. 今後の拡張ポイント（優先順）

1. **RCM 連携**: Risk/Control に ID を発番し、別シートの RCM と相互リンク
2. **差分比較**: 年度間のフロー変更をハイライト（統制の追加/廃止が可視化される）
3. **監査証跡**: いつ誰がどのノードを編集したかの履歴を YAML に残す
4. **複数業務プロセスの横断**: 販売 / 購買 / 在庫 / 決算 など複数フローの束を
   1 プロジェクトとして扱う `Program` エンティティ
5. **自動レイアウト改善**: レーン跨ぎのエッジ交差を最小化するヒューリスティック

---

## まとめ

- **コア**は単純な Flow DSL。J-SOX 固有要素（R/C/帳票/IT区分）は任意拡張。
- **エージェント**はドメインモデルを直接いじらず、限定されたツール経由で編集。
- **レンダラ**は中間 IR を挟み、用途別（閲覧用 Mermaid / 監査用 drawio）に切替可能。
- まずは提示サンプルを動かす最小実装 → Mermaid 出力 → R/C 拡張の順で育てる。
