' 貼り付け先: フォームモジュール F_SMS抽出条件
'   （フォーム名が違う場合は読み替え。VBE でフォームのモジュールを開いて貼り付ける）
'
' ■ 前提となるデザインビューでの設定
'   - フォームのレコードソース: T_SMS抽出条件（1レコード運用・条件ID=1）
'   - 既定のビュー: 単票フォーム
'   - 追加の許可: いいえ ／ 削除の許可: いいえ
'   - 移動ボタン: いいえ ／ レコードセレクタ: いいえ ／ スクロールバー: なし
'   - 各条件項目のテキストボックスは「既存のフィールドの追加」からドラッグして連結
'   - 更新日のテキストボックス: 編集ロック=はい（手入力させない）
'
' ■ このコードがやること
'   - 保存されるタイミング（BeforeUpdate）で更新日に今日の日付を自動セット
'     （時刻なしの Date。テーブルに更新日フィールドが無い場合はこの処理ごと不要）
'   - 保存ボタン（btn保存）を置く場合のクリック処理（任意。無ければ削除してよい）

Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(Cancel As Integer)
    Me!更新日 = Date
End Sub

' 任意: 明示的な保存ボタンを置く場合（ボタン名: btn保存）
' 連結フォームはフォームを閉じるだけでも自動保存されるため、無くても動く
Private Sub btn保存_Click()
    If Me.Dirty Then
        Me.Dirty = False    ' この時点で BeforeUpdate が走り保存される
        MsgBox "抽出条件を保存しました。", vbInformation
    Else
        MsgBox "変更はありません。", vbInformation
    End If
End Sub
