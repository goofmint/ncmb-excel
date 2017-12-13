# NCMB for Excel

ExcelからNCMBを操作しやすくするクラスモジュールです。現在、以下の機能が提供されています。

- データストア
  - データ保存
  - データ更新
  - データ検索

## 使い方

### 初期化

```vb
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

```vb
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

### データクラスの検索

```vb
Dim dataClass As clsDataStore
Set dataClass = ncmb.dataStore("Data")

Dim dataItems() As clsDataItem
dataClass.
dataClass.equalTo "message", "Hello World"
dataClass.greaterThan "Integer", 10
Dim dataItems() As clsDataItem
dataItems = dataClass.fetchAll()
Debug.Print dataItems(0).val("Integer")
```

### データストアの更新

```vb
Dim dataItem As clsDataItem
Set dataItem = dataClass.newData
dataItem.Field "message", "Hello World"
If dataItem.Save() Then
  Debug.Print ("保存できました")
Else
  Debug.Print ("保存失敗")
End If

dataItem.Field "message", "Update!"
If dataItem.Save() Then
    Debug.Print ("更新されました")
Else
    Debug.Print ("更新失敗")
End If
```

## データストアの削除

```vb
Dim dataItem As clsDataItem
Set dataItem = dataClass.newData
dataItem.Field "message", "Hello World"
If dataItem.Save() Then
  Debug.Print ("保存できました")
Else
  Debug.Print ("保存失敗")
End If

If dataItem.Delete() Then
    Debug.Print ("削除されました")
Else
    Debug.Print ("削除失敗")
End If
```

## LICENSE

MIT License