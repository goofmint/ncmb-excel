# NCMB for Excel

ExcelからNCMBを操作しやすくするクラスモジュールです。

## 使い方

### 初期化

```
Dim ApplicationKey As String
Dim ClientKey As String
ApplicationKey = "YOUR_APPLICATION_KEY"
ClientKey = "YOUR_CLIENT_KEY"

Dim ncmb As clsNCMB
Set ncmb = New clsNCMB
ncmb.ApplicationKey = ApplicationKey
ncmb.ClientKey = ClientKey
```

### データクラスの保存

```
Dim dataClass As clsDataStore
Set dataClass = ncmb.dataStore("Data")

Dim dataItem As clsDataItem
Set dataItem = dataClass.newData
dataItem.Field "message", "Hello World"
If dataItem.Save() Then
    Debug.Print ("保存できました")
Else
    Debug.Print ("保存失敗")
End If
```

## LICENSE

MIT License