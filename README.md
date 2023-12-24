# WarlanderRecordLog
Warlander戦績記録用マクロ


## マクロが壊れた場合

保存されていない状態かつ、xlsmを閉じていない状態で、  
PCをシャットダウンするとファイルが破損する。

破損時は、  
ボタンに使われているActiveXコントロールが図形になるという現象が起きる。

その為、  
破損前のxlsmからボタンをコピーし、  
プロパティから、オブジェクト名に正しい名前に書き換える必要がある。  
※ActiveXコントロールのコピー時は、オブジェクト名まではコピーできない。


## オブジェクト名を置き換える際のメモ

    AddRow
    AddColumn
    Victory
    Defeat
    Draw

    DeleteRow
    DeleteCol

    Team_ComboBox
    PlayerRank_ComboBox
    SquadRank_ComboBox
    TotalValor_TextBox
    Confirm
