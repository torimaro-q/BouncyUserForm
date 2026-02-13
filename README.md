# **BouncyUserForm**
(minimal)
![screenshot](pic/Minimal.gif)

(full)
![screenshot](pic/full.gif)

- BouncyUserForm は、**Excel VBA の UserForm を物理オブジェクトとして扱い、
重力・反発・空気抵抗・ダメージ表現を伴うアニメーションを実現するモジュール**です。
フォームをぶん投げてストレスを解消できます。
フォームが画面内を跳ね回り、衝突時にはコントロールが破損（非表示）します。
- 本コードでPCやデータに異常・損害が発生しても、作成者は一切責任を取りません。自己責任で遊んでください。

---
- BouncyUserForm is modules that treats an Excel VBA UserForm as a physical object, enabling animations with gravity, bouncing, air resistance, and damage effects.
You can throw the form around to relieve stress.
The form bounces around the screen, and when it collides with something, its controls can break (become hidden).
- The creator assumes no responsibility for any issues or damage to your PC or data caused by this code.
Use it at your own risk.

# 🐧 特徴 / Features
### 🧠物理エンジン / Physics Engine
- 重力シミュレーション
- 反発係数による跳ね返り
- 空気抵抗（速度依存）
- 衝突時のダメージ計算
- コントロールの破損（ランダム非表示）
- UserForm をドラッグして投げると物理挙動開始
- 画面端を壁として扱う衝突判定
- イベント通知（Move / Crash / Break / Started / Stopped）
---
- Gravity simulation
- Bounce with restitution coefficient
- Air resistance (velocity‑dependent)
- Damage calculation on impact
- Random control “breakage” (hidden on crash)
- Throw the UserForm by dragging it
- Screen edges act as collision walls
- Event callbacks: Move / Crash / Break / Started / Stopped
---

# 🧩拡張機能 / Extensions
以下のような拡張やエフェクトを自由に追加できます。
- OpenGL によるエフェクト
- Excel シートへのリアルタイム描画
- Excel 散布図による軌跡ログ
- 操作用 UI（コントローラー）
- ボロノイ分割をベースにした亀裂表現
- カスタムエフェクト（ICFormPhysicsEf）
- カスタム拡張（ICFormPhysicsEx）
---
You can freely add optional extensions such as:
- OpenGL‑based visual effects
- Real‑time drawing onto an Excel worksheet
- Trajectory logging using Excel scatter charts
- A controller UI for manual operation
- Voronoi-based crack representation
- Custom effects (via ICFormPhysicsEf)
- Custom extensions (via ICFormPhysicsEx)


### 🧩拡張 / Extensions
|||
|---|---|
| Excel ロガー / Excel logger | ![screenshot](pic/ex/CFormPhysicsLogger.gif) |
|シートレンダラー / Worksheet renderer|![screenshot](pic/ex/CFormPhysicsWsRenderer.gif)|
|コントローラー UI / Controller UI|![screenshot](pic/ex/CFormPhysicsController.gif)|
|フォーム描画 / Form renderer|![screenshot](pic/ex/CFormPhysicsFmRenderer.gif)|

### 💥エフェクト(OpenGL) / Effects(OpenGL)
|||
|---|---|
|爆発（glExplosion）|![screenshot](pic/ef/glExplosion.gif)|
|衝撃波（glShockWave）|![screenshot](pic/ef/glShockWave.gif)|
|移動残光（glMoveTrail）|![screenshot](pic/ef/glMoveTrail.gif)|
|コントロール破損(glControlShatter)|![screenshot](pic/ef/glControlShatter.gif)|
|ダメージ表示（glHitNumber）|![screenshot](pic/ef/glHitNumber.gif)|
|ステータス表示（glStatusVisualizer）|![screenshot](pic/ef/glStatusVisualizer.gif)|

# 🐧使い方 / Usage
## 1. クラスモジュールを追加 / Add the class modules
- **拡張なし（最小構成）/ Minimal setup (no extensions)**
```
(必須 / required)
+ CFormPhysics.cls
+ ICFormPhysicsEx.cls
```
- **拡張あり（OpenGL 以外）/ With extensions (non‑OpenGL)**
```
(必須 / required)
+ CFormPhysics.cls
+ ICFormPhysicsEx.cls
(任意 / optional)
+ CFormPhysicsLogger.cls
+ CFormPhysicsWsRenderer.cls
+ CFormPhysicsFmRenderer.cls
+ CFormPhysicsController.frm/frx
```
- **OpenGL 拡張あり / With OpenGL extensions**
```
(必須 / required)
+ CFormPhysics.cls
+ ICFormPhysicsEx.cls
+ ICFormPhysicsEf.cls
+ CFormPhysicsGLEffector.frm/frx
+ GLH.bas
+ OpenGL.cls
(任意 / optional)
+ glExplosion.cls
+ glShockWave.cls
+ glMoveTrail.cls
+ glControlShatter.cls
+ glHitNumber.cls
+ glStatusVisualizer.cls
```

## 2. UserForm にコードを追加 / Add code to your UserForm
- **拡張なし（最小構成）/ Minimal setup (no extensions)**
```vb
Private engine As CFormPhysics
Private Sub UserForm_Initialize()
    Set engine = New CFormPhysics
    engine.Init Me
End Sub
Private Sub UserForm_Terminate()
    engine.Terminate
End Sub
```
- **拡張あり（OpenGL 以外）/ With extensions (non‑OpenGL)**
- 使いたい機能を第2引数のArrayに入れる
```vb
Private engine As CFormPhysics
Private Sub UserForm_Initialize()
    Set engine = New CFormPhysics
    engine.Init Me, Array(CFormPhysicsLogger, CFormPhysicsWsRenderer)
End Sub
Private Sub UserForm_Terminate()
    engine.Terminate
End Sub
```
- **OpenGL 拡張あり / With OpenGL extensions**
    - 引数2 : CFormPhysicsGLEffector
    - 引数3 : Crash時に発生するエフェクト
    - 引数4 : Move時に発生するエフェクト
    - 引数5 : コントロール破損エフェクト
    ---
    - Argument 2: CFormPhysicsGLEffector
    - Argument 3: Effect triggered during a crash
    - Argument 4: Effect triggered during movement
    - Argument 5: Control‑shatter effect
```vb
Private engine As CFormPhysics
Private Sub UserForm_Initialize()
    Set engine = New CFormPhysics
    engine.init Me, Array(CFormPhysicsController), _
                    Array(glShockWave, _
                          glExplosion), _
                    Array(glMoveTrail), _
                    Array(glControlShatter)

End Sub
Private Sub UserForm_Terminate()
    engine.Terminate
End Sub
```
## 3. UserForm を表示 / Run the UserForm
- フォームをドラッグして投げると物理シミュレーションが開始します。
  ※タイトルバーではなく、ユーザーフォーム本体をドラッグしてください。
- Drag the form body (not the title bar) and release it to start the physics simulation.

# 🐧Requirements
- Windows + Excel (32‑bit / 64‑bit), likely Excel 2011 or later
- OpenGL (included with Windows)

# 🐧ライセンス / License
MIT License

# 🐧デモファイル / Demo file
[Sample](src/examples/DEMO.xlsm)

# 🔍 Search keywords
Excel physics engine, VBA animation, Excel game engine,
UserForm animation, OpenGL in VBA, Excel OpenGL, VBA graphics