

# Put your default configuration here
$PortName = "COM3"
$PortBaudRate = 115200
$PortParityBit = [System.IO.Ports.Parity]::None
$PortDataBit = 8
$PortStopBit = 1

# LF
$LFEndLine = "`n"

# CR-LF
$CRLFEndLine = "`r`n"


# Create a Serial Port Object
$SerialPort = new-Object System.IO.Ports.SerialPort $PortName, $PortBaudRate, $PortParityBit, $PortDataBit, $PortStopBit

# refer https://qiita.com/yapg57kon/items/58d7f47022b3e405b5f3

# 1. DTR、RTSを設定 (機器・ケーブル仕様に合わせる)
$SerialPort.DtrEnable = $true
$SerialPort.RtsEnable = $true

# 2. ハンドシェイク無し (必要なら::RequestToSendや::XOnXOffとする)
$SerialPort.Handshake = [System.IO.Ports.Handshake]::None

# 3. 改行文字をCR(0x0D)に設定 (機器仕様に合わせる)
#  ※NewLineプロパティはWriteLineやReadLineメソッドに適用されるため本方法では動作に影響しない
#  実際の変更方法の例は後述
$SerialPort.NewLine = "`r"

# 4. 文字コードをSJISに設定 (機器仕様に合わせる)
# $c.Encoding=[System.Text.Encoding]::GetEncoding("Shift_JIS")

# 5. シリアル受信イベントを登録(受信したらコンソールに出力)
$OnDataReceivedEventHandler = Register-ObjectEvent -InputObject $SerialPort -EventName "DataReceived" `
  -Action {param([System.IO.Ports.SerialPort]$sender, [System.EventArgs]$e) `
  Write-Host -NoNewline ($sender.ReadExisting()).Replace($LFEndLine, $CRLFEndLine)}


# 6. COMポートを開く
$SerialPort.Open()

Write-Host "Connected to SerialPort ${PortName} in ${PortBaudRate}Hz with <Parity, Data, Stop>:Bit = <$PortParityBit, $PortDataBit, $PortStopBit>"

Write-Host $SerialPort.ReadExisting()

# 7. キーボード入力をシリアルポートに送信する無限ループ (ReadKey($false)とすればローカルエコーになる)
#    これ以降、ターミナルソフトのようなキーボード入力と機器の出力表示になります (コピペデータも機器に送られるので注意)
#    終了時はctrl-cで抜ける
for(;;){if([Console]::KeyAvailable){$SerialPort.Write( `
    $(switch(([Console]::ReadKey($true)).KeyChar){{$_-eq[char]13}{$LFEndLine}default{$_}}))}}

# 9. COMポート閉じる
$SerialPort.Close()

# 10. イベント登録解除
Unregister-Event $OnDataReceivedEventHandler.Name

# 11. ジョブ削除
Remove-Job $OnDataReceivedEventHandler.id

