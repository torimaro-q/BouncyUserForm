# BouncyUserForm
- BouncyUserForm は、**Excel VBA の UserForm を物理オブジェクトとして扱い、
重力・反発・空気抵抗・ダメージ表現を伴うアニメーションを実現するクラスモジュール**です。
業務用のフォームに少なくとも1つのクラスを追加し、フォーム側へ数行追記するだけで、フォームをぶん投げてストレスを解消できます。
フォームが画面内を跳ね回り、衝突時にはコントロールが破損（非表示）します。
※本コードでPCやデータに異常・損害が発生しても、作成者は一切責任を取りません。自己責任で遊んでください。

- BouncyUserForm is a class module that treats an Excel VBA UserForm as a physical object, enabling animations with gravity, bounce, air resistance, and damage effects.
By adding at least one class to your production UserForm and inserting just a few lines of code, you can literally throw the form around to relieve stress.
The form will bounce around inside the screen, and controls will “break” (become hidden) upon impact.
Please note that the creator assumes no responsibility for any issues or damage to your PC or data caused by this code. Use it at your own risk and have fun.


![screenshot](pic/dentaku.gif)

## 🚀 特徴 / Features

- **重力シミュレーション**
- **反発係数による跳ね返り**
- **空気抵抗（速度依存）**
- **衝突時のダメージ計算**
- **ダメージに応じた背景色変化**
- **コントロールの破損（ランダム非表示）**
- **UserForm をドラッグして投げると物理挙動開始**
- **画面端を壁として扱う衝突判定**

---

- Gravity and bouncing 
- Air resistance 
- Damage calculation 
- Background color changes with damage 
- Controls randomly hide on impact 
- Throw the form by dragging it 
- Screen edges act as walls




## 📦 使い方

### 1. クラスモジュールを追加 / Add the class module 
- `CFormPhysics.cls`として本リポジトリのコードを貼り付けます。
- Create a class named **`CFormPhysics`** and paste the code.


### 2. UserForm に以下を追加 / Add this to your UserForm

```vb
Private engine As New CFormPhysics
Private Sub UserForm_Activate()
    engine.Init Me
End Sub
```

### 3. UserForm を表示 / Run the form

- フォームをドラッグして投げると物理シミュレーションが開始します。
  
  ※タイトルバーではなく、ユーザーフォーム本体をドラッグしてください。
- Run the form and throw it to start the physics.
  
  Note: Drag the form itself, not the title bar.



### ライセンス / License
MIT License

