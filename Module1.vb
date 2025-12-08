Module Module1
    '
    'Ver1.10  R側にアラームコメントを入れるとXbar側にアラームが出てしまう不具合修正
    '         アラームコメント入力後QC承認待ちの場合、ツリーアイコンの表示をQマークに変更。またOK MONITORの表示も赤点滅しないように変更。

    'Ver1.12 管理値変更時のエラー修正

    'Ver1.13 MR管理図(変動管理図)追加。
    '二重起動防止機能追加

    'Ver1.16 管理値新規設定機能の不具合修正

    'Ver1.17 プロダクトモニターコンバータを追加
    '        MR管理図を適用している測定項目の登録時にエラーになる不具合を修正

    'Ver1.18 管理値・SPCルール変更時の証拠残し機能を追加
    '        MR管理図の管理値登録フォームを追加
    '        管理限界変化率CLCR、OOCを算出する機能を追加

    'Ver1.20 プロダクトモニタとのソケット連動機能を追加
    '        管理値登録時にQC承認がされていないと登録できない不具合を修正

    'Ver1.21 CPK計算機能を追加
    '        測定が1点のみでRが無い場合は、Rの表示領域をグレーで塗りつぶす機能を追加。

    'Ver1.22 画面表示範囲のCpkを表示する機能を追加
    '    
    'Ver1.26 コーターPMS追加
    '    
    'Ver1.28 画面解像度1920×1080,1280×1024,1024×768の三パターンの画面を作成し、画面サイズを切り替える機能を追加
    '日本語表示と英語表示を切り替える機能を追加
    'ユーザーマニュアルを作成し、ソフトから閲覧できる機能を追加

    'Ver1.32
    'データの更新方法を変更

    'Config用変数
    Public StrCDir As String
    Public StrRootFolder As String
    Public StrNetworkFolder As String
    Public MajorItem As String
    Public MyHostName As String
    Public ServerName As String
    Public StrLanguage As String '翻訳言語
    Public StrResolution As String '解像度

    'グラフ作成用変数
    Public X_USL, X_LSL, X_SCL, X_kousa As Double 'XBar側規格値
    Public R_USL, R_LSL As Double 'R側規格値
    Public X_UCL, X_LCL, X_CL, X_Shiguma As Double 'XBar側管理値
    Public R_UCL, R_LCL, R_CL, R_Shiguma As Double 'R側管理値
    Public MR_UCL, MR_LCL, MR_CL, MR_Shiguma As Double 'MR側管理値
    Public X_gType As String
    Public SPCDataNum As Integer
    Public DispStartPosition As Integer
    Public ycl, yucl, ylcl As Integer
    Public xpnbuf_X(5000) As Integer
    Public ypnbuf_X(5000) As Integer
    Public xpnbuf_R(5000) As Integer
    Public ypnbuf_R(5000) As Integer
    Public Graphsmallcount As Integer
    Public strUnit As String ' 単位
  


    'グラフデータ用バッファー
    Public vv As Integer = 10000
    Public MesureValueBuf(vv) As String '測定値


    Public strAlarmStartDate As Date
    Public strStartDate As Date

    'SPCアラーム用
    Public QCNotCheckFlag As Boolean
    Public HostSub As String


    Public SerectPoint As Integer

    Public PropertyNo As Integer

    Public MessageBoxShowFlag As Boolean

    Public MRFlag As Boolean

    Public StrServerConnection As String = "Server=MA8104520;User ID=sa;Password=1234;Initial Catalog=EQP_MONITOR"
    'Public StrServerConnection As String = "Server=10.22.34.19;User ID=sa;Password=Lsipm1234;Initial Catalog=EQP_MONITOR"
    Public StrErrMes As String
    Public PropertyTable As New DataTable
    Public TreeName() As String

    Public UserName As String 'パスワード照合用
    Public JP_Message As String 'パスワード照合用
    Public EN_Message As String 'パスワード照合用

    Public gType(3 - 1) As String

    Public M_Data() As String
    Public _id As Integer = 0
    Public _wDate As Integer = 1
    Public _X As Integer = 2
    Public _R As Integer = 3
    Public _MR As Integer = 4
    Public _opName As Integer = 5
    Public _lot As Integer = 6
    Public _cate As Integer = 7
    Public M_Alarm()() As String
    '****************************************************************************************************************************************
    '       ログファイル出力
    '****************************************************************************************************************************************
    Public Sub SaveLog(ByVal NowTime As String, ByVal ErrMes As String)

        Dim FileName As String = StrCDir & "\Log.log"
        Dim LogText As String

        Try
            LogText = NowTime + ", " + ErrMes
            Dim ErrLogFile As New System.IO.StreamWriter(FileName, True, System.Text.Encoding.GetEncoding("Shift_Jis"))
            ErrLogFile.WriteLine(LogText)
            ErrLogFile.Close()
            Exit Sub
        Catch ex As Exception
            Exit Sub
        End Try

    End Sub


End Module
