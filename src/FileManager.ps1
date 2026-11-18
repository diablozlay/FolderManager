#requires -Version 5.1
<#
    檔案管理工具 Folder Manager
    - 列出所有可用磁碟機（可多選）
    - 依關鍵字掃描資料夾（可選檔案），支援深度限制與中斷
    - 「立即停止」按鈕 / Ctrl+Q：隨時中止所有作業並結束程式
    - 勾選後刪除：預設「移至回收筒」（可復原），可改永久刪除
    - 刪除前確認視窗，列出所有將被刪除的項目
    - 全程寫入 logs\operations_yyyyMMdd.log
    - 權限不足 / 檔案被佔用 / 路徑過長 等例外逐項記錄並繼續
#>

param([switch]$SelfTest, [switch]$Perf)
$script:PerfMode = [bool]$Perf

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$script:RecycleOk = $true
try { Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop } catch { $script:RecycleOk = $false }
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------- 日誌
# 可攜性：優先使用啟動器傳入的 FM_BASEDIR（.exe 所在資料夾），
# 其次 $PSScriptRoot（.ps1 模式），最後才是目前工作目錄。
$script:BaseDir = $env:FM_BASEDIR
if ([string]::IsNullOrWhiteSpace($script:BaseDir)) { $script:BaseDir = $PSScriptRoot }
if ([string]::IsNullOrWhiteSpace($script:BaseDir)) { $script:BaseDir = (Get-Location).Path }

$script:LogDir = Join-Path $script:BaseDir 'logs'
try {
    if (-not (Test-Path -LiteralPath $script:LogDir)) {
        New-Item -ItemType Directory -Path $script:LogDir -Force -ErrorAction Stop | Out-Null
    }
} catch {
    # 資料夾不可寫（例如放在唯讀目錄）→ 退回使用者設定檔
    $script:LogDir = Join-Path $env:LOCALAPPDATA 'FolderManager\logs'
    if (-not (Test-Path -LiteralPath $script:LogDir)) {
        New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null
    }
}
$script:LogFile = Join-Path $script:LogDir ('operations_{0}.log' -f (Get-Date -Format 'yyyyMMdd'))
$script:Utf8 = New-Object System.Text.UTF8Encoding($true)
$script:LogBox = $null
# 批次日誌：檔案與畫面各用一個緩衝，避免每列都開關檔案、重排 TextBox。
# 一般模式每列立即落盤；大量刪除時由刪除迴圈開啟 LogBatching，每 100 列或 250ms 落盤一次。
$script:LogLineCount = 0
$script:LogFilePending = New-Object System.Text.StringBuilder
$script:LogUiPending = New-Object System.Text.StringBuilder
$script:LogBatching = $false
$script:LogClock = [System.Diagnostics.Stopwatch]::StartNew()

# ---------------------------------------------------------------- 語言 / 在地化
# 語言碼取自 ZW3D 安裝目錄 languages\LanguageMap.txt；Name 為該語言自稱，下拉選項固定不翻譯
$script:Langs = @(
    @{ Id='cs_CZ'; Name='Čeština';    Font='Segoe UI' }
    @{ Id='de_DE'; Name='Deutsch';    Font='Segoe UI' }
    @{ Id='en_US'; Name='English';    Font='Segoe UI' }
    @{ Id='es_ES'; Name='Español';    Font='Segoe UI' }
    @{ Id='fr_FR'; Name='Français';   Font='Segoe UI' }
    @{ Id='hu_HU'; Name='Magyar';     Font='Segoe UI' }
    @{ Id='it_IT'; Name='Italiano';   Font='Segoe UI' }
    @{ Id='ja_JP'; Name='日本語';      Font='Yu Gothic UI' }
    @{ Id='ko_KR'; Name='한국어';       Font='Malgun Gothic' }
    @{ Id='pl_PL'; Name='Polski';     Font='Segoe UI' }
    @{ Id='pt_PT'; Name='Português';  Font='Segoe UI' }
    @{ Id='ru_RU'; Name='Русский';    Font='Segoe UI' }
    @{ Id='tr_TR'; Name='Türkçe';     Font='Segoe UI' }
    @{ Id='zh_CN'; Name='中文(简体)';  Font='Microsoft YaHei UI' }
    @{ Id='zh_TW'; Name='中文(繁體)';  Font='Microsoft JhengHei UI' }
)
$script:LangFile = Join-Path $script:BaseDir 'language.txt'
$script:ScanLangId = ''             # '' = 全部語言；否則為單一 ZW3D 語言版本（僅影響掃描關鍵詞）
$script:UiLang = 'zh_TW'            # 本工具介面語言，預設繁中；與掃描語言版本完全獨立
$script:ApplyingLang = $false
$script:IconButtons = New-Object System.Collections.ArrayList

# 詞條表：以 key 為主鍵。未提供的語言依序回退 en_US → zh_TW。
# 繁中為原版文案；簡中／英文完整；其餘語言提供主要按鈕與短標籤，長提示回退英文。
$script:Str = @{
 'AppTitle'      = @{ zh_TW='檔案管理工具 · Folder Manager'; zh_CN='文件管理工具 · Folder Manager'; en_US='Folder Manager · ZW3D Files'; de_DE='Dateiverwaltung · ZW3D'; fr_FR='Gestionnaire de fichiers · ZW3D'; es_ES='Gestor de archivos · ZW3D'; it_IT='Gestione file · ZW3D'; pt_PT='Gestor de ficheiros · ZW3D'; ja_JP='ファイル管理ツール · ZW3D'; ko_KR='파일 관리자 · ZW3D'; ru_RU='Менеджер файлов · ZW3D'; pl_PL='Menedżer plików · ZW3D'; cs_CZ='Správce souborů · ZW3D'; hu_HU='Fájlkezelő · ZW3D'; tr_TR='Dosya Yöneticisi · ZW3D' }
 'KwLabel'       = @{ zh_TW='檔名／資料夾名稱關鍵字'; zh_CN='文件名／文件夹名关键字'; en_US='File / folder name keyword'; de_DE='Datei- / Ordnername'; fr_FR='Mot-clé (fichier / dossier)'; es_ES='Palabra clave (archivo / carpeta)'; it_IT='Parola chiave (file / cartella)'; pt_PT='Palavra-chave (ficheiro / pasta)'; ja_JP='ファイル／フォルダー名'; ko_KR='파일/폴더 이름 키워드'; ru_RU='Ключевое слово (файл / папка)'; pl_PL='Słowo kluczowe (plik / folder)'; cs_CZ='Klíčové slovo (soubor / složka)'; hu_HU='Kulcsszó (fájl / mappa)'; tr_TR='Anahtar kelime (dosya / klasör)' }
 'PresetLabel'   = @{ zh_TW='常用'; zh_CN='常用'; en_US='Preset'; de_DE='Vorlage'; fr_FR='Favoris'; es_ES='Frecuentes'; it_IT='Predefiniti'; pt_PT='Predefinições'; ja_JP='よく使う'; ko_KR='자주 쓰는'; ru_RU='Шаблон'; pl_PL='Szablon'; cs_CZ='Šablona'; hu_HU='Sablon'; tr_TR='Hazır' }
 'TargetLabel'   = @{ zh_TW='搜尋對象'; zh_CN='搜索对象'; en_US='Search for'; de_DE='Suchen nach'; fr_FR='Rechercher'; es_ES='Buscar'; it_IT='Cerca'; pt_PT='Pesquisar'; ja_JP='検索対象'; ko_KR='검색 대상'; ru_RU='Искать'; pl_PL='Szukaj'; cs_CZ='Hledat'; hu_HU='Keresés'; tr_TR='Ara' }
 'DepthLabel'    = @{ zh_TW='深度'; zh_CN='深度'; en_US='Depth'; de_DE='Tiefe'; fr_FR='Profondeur'; es_ES='Profundidad'; it_IT='Profondità'; pt_PT='Profundidade'; ja_JP='深さ'; ko_KR='깊이'; ru_RU='Глубина'; pl_PL='Głębokość'; cs_CZ='Hloubka'; hu_HU='Mélység'; tr_TR='Derinlik' }
 'PermanentChk'  = @{ zh_TW='永久刪除（不進回收筒）'; zh_CN='永久删除（不进回收站）'; en_US='Permanent delete (skip recycle bin)'; de_DE='Endgültig löschen (ohne Papierkorb)'; fr_FR='Suppression définitive (sans corbeille)'; es_ES='Eliminar permanentemente (sin papelera)'; it_IT='Eliminazione definitiva (senza cestino)'; pt_PT='Eliminar definitivamente (sem reciclagem)'; ja_JP='完全に削除（ごみ箱不使用）'; ko_KR='영구 삭제 (휴지통 사용 안 함)'; ru_RU='Удалить безвозвратно (минуя корзину)'; pl_PL='Usuń trwale (bez kosza)'; cs_CZ='Trvale smazat (bez koše)'; hu_HU='Végleges törlés (lomtár nélkül)'; tr_TR='Kalıcı sil (Geri Dönüşüm Kutusu olmadan)' }
 'Hint'          = @{ zh_TW='提示：關鍵字留空 = 列出全部（受結果上限 5000 筆限制）。Enter 直接搜尋。'; zh_CN='提示：关键字留空 = 列出全部（受结果上限 5000 条限制）。Enter 直接搜索。'; en_US='Hint: leave the keyword empty to list all (max 5000 results). Press Enter to search.' }
 'Zw3dOnlyChk'   = @{ zh_TW='僅顯示 ZW3D 相關'; zh_CN='仅显示 ZW3D 相关'; en_US='ZW3D files only'; de_DE='Nur ZW3D-Dateien'; fr_FR='Fichiers ZW3D uniquement'; es_ES='Solo archivos ZW3D'; it_IT='Solo file ZW3D'; pt_PT='Apenas ficheiros ZW3D'; ja_JP='ZW3D 関連のみ'; ko_KR='ZW3D 관련만'; ru_RU='Только файлы ZW3D'; pl_PL='Tylko pliki ZW3D'; cs_CZ='Pouze soubory ZW3D'; hu_HU='Csak ZW3D fájlok'; tr_TR='Yalnızca ZW3D dosyaları' }
 'ExtLabel'      = @{ zh_TW='副檔名'; zh_CN='扩展名'; en_US='Extension'; de_DE='Erweiterung'; fr_FR='Extension'; es_ES='Extensión'; it_IT='Estensione'; pt_PT='Extensão'; ja_JP='拡張子'; ko_KR='확장자'; ru_RU='Расширение'; pl_PL='Rozszerzenie'; cs_CZ='Přípona'; hu_HU='Kiterjesztés'; tr_TR='Uzantı' }
 'Zw3dHint'      = @{ zh_TW='判斷：副檔名清單／檔名含 ZW3D·TopologicRep·Proxy／路徑含 ZW3D'; zh_CN='判断：扩展名清单／文件名含 ZW3D·TopologicRep·Proxy／路径含 ZW3D'; en_US='Match: extension list / name contains ZW3D·TopologicRep·Proxy / path contains ZW3D' }
 'UiLangLabel'   = @{ zh_TW='介面語言'; zh_CN='界面语言'; en_US='Language'; de_DE='Sprache'; fr_FR='Langue'; es_ES='Idioma'; it_IT='Lingua'; pt_PT='Idioma'; ja_JP='表示言語'; ko_KR='표시 언어'; ru_RU='Язык'; pl_PL='Język'; cs_CZ='Jazyk'; hu_HU='Nyelv'; tr_TR='Dil' }
 'LangLabel'     = @{ zh_TW='ZW3D 語言版本'; zh_CN='ZW3D 语言版本'; en_US='ZW3D language'; de_DE='ZW3D Sprache'; fr_FR='Langue ZW3D'; es_ES='Idioma ZW3D'; it_IT='Lingua ZW3D'; pt_PT='Idioma ZW3D'; ja_JP='ZW3D 言語'; ko_KR='ZW3D 언어'; ru_RU='Язык ZW3D'; pl_PL='Język ZW3D'; cs_CZ='Jazyk ZW3D'; hu_HU='ZW3D nyelv'; tr_TR='ZW3D dili' }
 'Target.Folder' = @{ zh_TW='資料夾'; zh_CN='文件夹'; en_US='Folders'; de_DE='Ordner'; fr_FR='Dossiers'; es_ES='Carpetas'; it_IT='Cartelle'; pt_PT='Pastas'; ja_JP='フォルダー'; ko_KR='폴더'; ru_RU='Папки'; pl_PL='Foldery'; cs_CZ='Složky'; hu_HU='Mappák'; tr_TR='Klasörler' }
 'Target.File'   = @{ zh_TW='檔案'; zh_CN='文件'; en_US='Files'; de_DE='Dateien'; fr_FR='Fichiers'; es_ES='Archivos'; it_IT='File'; pt_PT='Ficheiros'; ja_JP='ファイル'; ko_KR='파일'; ru_RU='Файлы'; pl_PL='Pliki'; cs_CZ='Soubory'; hu_HU='Fájlok'; tr_TR='Dosyalar' }
 'Target.Both'   = @{ zh_TW='兩者'; zh_CN='两者'; en_US='Both'; de_DE='Beide'; fr_FR='Les deux'; es_ES='Ambos'; it_IT='Entrambi'; pt_PT='Ambos'; ja_JP='両方'; ko_KR='둘 다'; ru_RU='Все'; pl_PL='Oba'; cs_CZ='Obojí'; hu_HU='Mindkettő'; tr_TR='İkisi de' }
 'Depth.3'       = @{ zh_TW='3 層'; zh_CN='3 层'; en_US='3 levels'; de_DE='3 Ebenen'; fr_FR='3 niveaux'; es_ES='3 niveles'; it_IT='3 livelli'; pt_PT='3 níveis'; ja_JP='3 階層'; ko_KR='3 단계'; ru_RU='3 уровня'; pl_PL='3 poziomy'; cs_CZ='3 úrovně'; hu_HU='3 szint'; tr_TR='3 seviye' }
 'Depth.5'       = @{ zh_TW='5 層'; zh_CN='5 层'; en_US='5 levels'; de_DE='5 Ebenen'; fr_FR='5 niveaux'; es_ES='5 niveles'; it_IT='5 livelli'; pt_PT='5 níveis'; ja_JP='5 階層'; ko_KR='5 단계'; ru_RU='5 уровней'; pl_PL='5 poziomów'; cs_CZ='5 úrovní'; hu_HU='5 szint'; tr_TR='5 seviye' }
 'Depth.8'       = @{ zh_TW='8 層'; zh_CN='8 层'; en_US='8 levels'; de_DE='8 Ebenen'; fr_FR='8 niveaux'; es_ES='8 niveles'; it_IT='8 livelli'; pt_PT='8 níveis'; ja_JP='8 階層'; ko_KR='8 단계'; ru_RU='8 уровней'; pl_PL='8 poziomów'; cs_CZ='8 úrovní'; hu_HU='8 szint'; tr_TR='8 seviye' }
 'Depth.12'      = @{ zh_TW='12 層'; zh_CN='12 层'; en_US='12 levels'; de_DE='12 Ebenen'; fr_FR='12 niveaux'; es_ES='12 niveles'; it_IT='12 livelli'; pt_PT='12 níveis'; ja_JP='12 階層'; ko_KR='12 단계'; ru_RU='12 уровней'; pl_PL='12 poziomów'; cs_CZ='12 úrovní'; hu_HU='12 szint'; tr_TR='12 seviye' }
 'Depth.Unlimited'= @{ zh_TW='不限'; zh_CN='不限'; en_US='Unlimited'; de_DE='Unbegrenzt'; fr_FR='Illimité'; es_ES='Sin límite'; it_IT='Illimitato'; pt_PT='Sem limite'; ja_JP='無制限'; ko_KR='제한 없음'; ru_RU='Без ограничений'; pl_PL='Bez limitu'; cs_CZ='Bez omezení'; hu_HU='Korlátlan'; tr_TR='Sınırsız' }
 'Ext.All'       = @{ zh_TW='（不限）'; zh_CN='（不限）'; en_US='(all)'; de_DE='(alle)'; fr_FR='(tous)'; es_ES='(todos)'; it_IT='(tutti)'; pt_PT='(todos)'; ja_JP='(すべて)'; ko_KR='(전체)'; ru_RU='(все)'; pl_PL='(wszystkie)'; cs_CZ='(vše)'; hu_HU='(mind)'; tr_TR='(tümü)' }
 'Preset.All'    = @{ zh_TW='（全部）'; zh_CN='（全部）'; en_US='(all)'; de_DE='(alle)'; fr_FR='(tous)'; es_ES='(todos)'; it_IT='(tutti)'; pt_PT='(todos)'; ja_JP='(すべて)'; ko_KR='(전체)'; ru_RU='(все)'; pl_PL='(wszystkie)'; cs_CZ='(vše)'; hu_HU='(mind)'; tr_TR='(tümü)' }
 'ZwScope.All'   = @{ zh_TW='全部 ZW3D 檔案（模型＋暫存）'; zh_CN='全部 ZW3D 文件（模型＋暂存）'; en_US='All ZW3D files (models + temp)' }
 'ZwScope.Junk'  = @{ zh_TW='僅 ZW3D 暫存垃圾檔（對應 clear ALL.bat）'; zh_CN='仅 ZW3D 暂存垃圾文件（对应 clear ALL.bat）'; en_US='ZW3D temp junk only (clear ALL.bat)' }
 'ZwScope.Model' = @{ zh_TW='僅 ZW3D 模型檔（.z3/.z3prt/.z3asm/.z3drw）'; zh_CN='仅 ZW3D 模型文件（.z3/.z3prt/.z3asm/.z3drw）'; en_US='ZW3D model files only (.z3/.z3prt/.z3asm/.z3drw)' }
 'Col.Name'      = @{ zh_TW='名稱'; zh_CN='名称'; en_US='Name'; de_DE='Name'; fr_FR='Nom'; es_ES='Nombre'; it_IT='Nome'; pt_PT='Nome'; ja_JP='名前'; ko_KR='이름'; ru_RU='Имя'; pl_PL='Nazwa'; cs_CZ='Název'; hu_HU='Név'; tr_TR='Ad' }
 'Col.Type'      = @{ zh_TW='類型'; zh_CN='类型'; en_US='Type'; de_DE='Typ'; fr_FR='Type'; es_ES='Tipo'; it_IT='Tipo'; pt_PT='Tipo'; ja_JP='種類'; ko_KR='유형'; ru_RU='Тип'; pl_PL='Typ'; cs_CZ='Typ'; hu_HU='Típus'; tr_TR='Tür' }
 'Col.Path'      = @{ zh_TW='完整路徑'; zh_CN='完整路径'; en_US='Full path'; de_DE='Vollständiger Pfad'; fr_FR='Chemin complet'; es_ES='Ruta completa'; it_IT='Percorso completo'; pt_PT='Caminho completo'; ja_JP='完全パス'; ko_KR='전체 경로'; ru_RU='Полный путь'; pl_PL='Pełna ścieżka'; cs_CZ='Úplná cesta'; hu_HU='Teljes útvonal'; tr_TR='Tam yol' }
 'Col.Parent'    = @{ zh_TW='上層資料夾'; zh_CN='上层文件夹'; en_US='Parent folder'; de_DE='Übergeordneter Ordner'; fr_FR='Dossier parent'; es_ES='Carpeta superior'; it_IT='Cartella padre'; pt_PT='Pasta principal'; ja_JP='親フォルダー'; ko_KR='상위 폴더'; ru_RU='Родительская папка'; pl_PL='Folder nadrzędny'; cs_CZ='Nadřazená složka'; hu_HU='Szülőmappa'; tr_TR='Üst klasör' }
 'Type.File'     = @{ zh_TW='檔案'; zh_CN='文件'; en_US='File'; de_DE='Datei'; fr_FR='Fichier'; es_ES='Archivo'; it_IT='File'; pt_PT='Ficheiro'; ja_JP='ファイル'; ko_KR='파일'; ru_RU='Файл'; pl_PL='Plik'; cs_CZ='Soubor'; hu_HU='Fájl'; tr_TR='Dosya' }
 'Type.Folder'   = @{ zh_TW='資料夾'; zh_CN='文件夹'; en_US='Folder'; de_DE='Ordner'; fr_FR='Dossier'; es_ES='Carpeta'; it_IT='Cartella'; pt_PT='Pasta'; ja_JP='フォルダー'; ko_KR='폴더'; ru_RU='Папка'; pl_PL='Folder'; cs_CZ='Složka'; hu_HU='Mappa'; tr_TR='Klasör' }
 'Btn.Search'    = @{ zh_TW='搜尋'; zh_CN='搜索'; en_US='Search'; de_DE='Suchen'; fr_FR='Rechercher'; es_ES='Buscar'; it_IT='Cerca'; pt_PT='Pesquisar'; ja_JP='検索'; ko_KR='검색'; ru_RU='Поиск'; pl_PL='Szukaj'; cs_CZ='Hledat'; hu_HU='Keresés'; tr_TR='Ara' }
 'Btn.Stop'      = @{ zh_TW='停止'; zh_CN='停止'; en_US='Stop'; de_DE='Stopp'; fr_FR='Arrêter'; es_ES='Detener'; it_IT='Ferma'; pt_PT='Parar'; ja_JP='停止'; ko_KR='중지'; ru_RU='Стоп'; pl_PL='Zatrzymaj'; cs_CZ='Zastavit'; hu_HU='Leállít'; tr_TR='Durdur' }
 'Btn.CheckAll'  = @{ zh_TW='勾選全部'; zh_CN='全选'; en_US='Select all'; de_DE='Alle auswählen'; fr_FR='Tout sélectionner'; es_ES='Seleccionar todo'; it_IT='Seleziona tutto'; pt_PT='Selecionar tudo'; ja_JP='すべて選択'; ko_KR='모두 선택'; ru_RU='Выбрать все'; pl_PL='Zaznacz wszystko'; cs_CZ='Vybrat vše'; hu_HU='Mind kijelöl'; tr_TR='Tümünü seç' }
 'Btn.Uncheck'   = @{ zh_TW='清除勾選'; zh_CN='清除勾选'; en_US='Clear selection'; de_DE='Auswahl aufheben'; fr_FR='Effacer la sélection'; es_ES='Borrar selección'; it_IT='Deseleziona tutto'; pt_PT='Limpar seleção'; ja_JP='選択解除'; ko_KR='선택 해제'; ru_RU='Снять выделение'; pl_PL='Wyczyść zaznaczenie'; cs_CZ='Zrušit výběr'; hu_HU='Kijelölés törlése'; tr_TR='Seçimi temizle' }
 'Btn.Delete'    = @{ zh_TW='刪除勾選項目'; zh_CN='删除勾选项'; en_US='Delete selected'; de_DE='Auswahl löschen'; fr_FR='Supprimer la sélection'; es_ES='Eliminar selección'; it_IT='Elimina selezionati'; pt_PT='Eliminar seleção'; ja_JP='選択項目を削除'; ko_KR='선택 항목 삭제'; ru_RU='Удалить выбранные'; pl_PL='Usuń zaznaczone'; cs_CZ='Smazat vybrané'; hu_HU='Kijelöltek törlése'; tr_TR='Seçilenleri sil' }
 'Btn.Explorer'  = @{ zh_TW='檔案總管'; zh_CN='文件资源管理器'; en_US='Explorer'; de_DE='Explorer'; fr_FR='Explorateur'; es_ES='Explorador'; it_IT='Esplora risorse'; pt_PT='Explorador'; ja_JP='エクスプローラー'; ko_KR='탐색기'; ru_RU='Проводник'; pl_PL='Eksplorator'; cs_CZ='Průzkumník'; hu_HU='Intéző'; tr_TR='Dosya Gezgini' }
 'Btn.LogFolder' = @{ zh_TW='日誌資料夾'; zh_CN='日志文件夹'; en_US='Log folder'; de_DE='Protokollordner'; fr_FR='Dossier journaux'; es_ES='Carpeta de registros'; it_IT='Cartella log'; pt_PT='Pasta de registos'; ja_JP='ログフォルダー'; ko_KR='로그 폴더'; ru_RU='Папка журналов'; pl_PL='Folder dzienników'; cs_CZ='Složka protokolů'; hu_HU='Naplómappa'; tr_TR='Günlük klasörü' }
 'Btn.Export'    = @{ zh_TW='匯出清單'; zh_CN='导出清单'; en_US='Export list'; de_DE='Liste exportieren'; fr_FR='Exporter la liste'; es_ES='Exportar lista'; it_IT='Esporta elenco'; pt_PT='Exportar lista'; ja_JP='一覧を出力'; ko_KR='목록 내보내기'; ru_RU='Экспорт списка'; pl_PL='Eksportuj listę'; cs_CZ='Exportovat seznam'; hu_HU='Lista exportálása'; tr_TR='Listeyi dışa aktar' }
 'Btn.Abort'     = @{ zh_TW='立即停止'; zh_CN='立即停止'; en_US='Abort'; de_DE='Sofort beenden'; fr_FR='Arrêt immédiat'; es_ES='Detención inmediata'; it_IT='Arresta immediato'; pt_PT='Parar agora'; ja_JP='即時終了'; ko_KR='즉시 중지'; ru_RU='Немедленный останов'; pl_PL='Natychmiast zatrzymaj'; cs_CZ='Okamžitě ukončit'; hu_HU='Azonnal leállít'; tr_TR='Hemen durdur' }
 'Btn.Refresh'   = @{ zh_TW='重新整理'; zh_CN='刷新'; en_US='Refresh'; de_DE='Aktualisieren'; fr_FR='Actualiser'; es_ES='Actualizar'; it_IT='Aggiorna'; pt_PT='Atualizar'; ja_JP='更新'; ko_KR='새로 고침'; ru_RU='Обновить'; pl_PL='Odśwież'; cs_CZ='Obnovit'; hu_HU='Frissít'; tr_TR='Yenile' }
 'Btn.ToggleDrives'= @{ zh_TW='全選/全不選'; zh_CN='全选/全不选'; en_US='Toggle all'; de_DE='Alle umschalten'; fr_FR='Inverser tout'; es_ES='Alternar todo'; it_IT='Inverti tutto'; pt_PT='Alternar tudo'; ja_JP='すべて切替'; ko_KR='전체 전환'; ru_RU='Переключить все'; pl_PL='Przełącz wszystko'; cs_CZ='Přepnout vše'; hu_HU='Mind átvált'; tr_TR='Tümünü değiştir' }
 'DrivesLabel'   = @{ zh_TW='磁碟機（可多選）'; zh_CN='磁盘（可多选）'; en_US='Drives (multi-select)'; de_DE='Laufwerke (Mehrfachauswahl)'; fr_FR='Disques (sélection multiple)'; es_ES='Unidades (selección múltiple)'; it_IT='Unità (selezione multipla)'; pt_PT='Unidades (seleção múltipla)'; ja_JP='ドライブ（複数選択可）'; ko_KR='드라이브 (다중 선택)'; ru_RU='Диски (множественный выбор)'; pl_PL='Dyski (wielokrotny wybór)'; cs_CZ='Disky (vícenásobný výběr)'; hu_HU='Meghajtók (többszörös kijelölés)'; tr_TR='Sürücüler (çoklu seçim)' }
 'Status.Ready'  = @{ zh_TW='就緒'; zh_CN='就绪'; en_US='Ready'; de_DE='Bereit'; fr_FR='Prêt'; es_ES='Listo'; it_IT='Pronto'; pt_PT='Pronto'; ja_JP='準備完了'; ko_KR='준비'; ru_RU='Готово'; pl_PL='Gotowe'; cs_CZ='Připraveno'; hu_HU='Kész'; tr_TR='Hazır' }
 'CreditTip'     = @{ zh_TW='開發者'; zh_CN='开发者'; en_US='Developer' }
 'Menu.Open'     = @{ zh_TW='在檔案總管中開啟'; zh_CN='在文件资源管理器中打开'; en_US='Open in Explorer'; de_DE='Im Explorer öffnen'; fr_FR="Ouvrir dans l'explorateur"; es_ES='Abrir en el Explorador'; it_IT='Apri in Esplora risorse'; pt_PT='Abrir no Explorador'; ja_JP='エクスプローラーで開く'; ko_KR='탐색기에서 열기'; ru_RU='Открыть в Проводнике'; pl_PL='Otwórz w Eksploratorze'; cs_CZ='Otevřít v Průzkumníku'; hu_HU='Megnyitás az Intézőben'; tr_TR='Gezgini aç' }
 'Menu.Copy'     = @{ zh_TW='複製完整路徑'; zh_CN='复制完整路径'; en_US='Copy full path'; de_DE='Vollständigen Pfad kopieren'; fr_FR='Copier le chemin complet'; es_ES='Copiar ruta completa'; it_IT='Copia percorso completo'; pt_PT='Copiar caminho completo'; ja_JP='完全パスをコピー'; ko_KR='전체 경로 복사'; ru_RU='Копировать полный путь'; pl_PL='Kopiuj pełną ścieżkę'; cs_CZ='Kopírovat úplnou cestu'; hu_HU='Teljes útvonal másolása'; tr_TR='Tam yolu kopyala' }
 'Menu.Delete'   = @{ zh_TW='刪除此項目'; zh_CN='删除此项'; en_US='Delete this item'; de_DE='Dieses Element löschen'; fr_FR='Supprimer cet élément'; es_ES='Eliminar este elemento'; it_IT='Elimina questo elemento'; pt_PT='Eliminar este item'; ja_JP='この項目を削除'; ko_KR='이 항목 삭제'; ru_RU='Удалить этот элемент'; pl_PL='Usuń ten element'; cs_CZ='Smazat tuto položku'; hu_HU='Elem törlése'; tr_TR='Bu öğeyi sil' }
 'Tip.Search'    = @{ zh_TW='開始搜尋（或在關鍵字欄位直接按 Enter）'; zh_CN='开始搜索（或在关键字栏直接按 Enter）'; en_US='Start search (or press Enter in the keyword box)' }
 'Tip.Stop'      = @{ zh_TW='停止目前進行中的掃描'; zh_CN='停止当前扫描'; en_US='Stop the running scan' }
 'Tip.CheckAll'  = @{ zh_TW='把清單中的每一筆結果都打上勾（準備刪除）'; zh_CN='勾选清单中的每一条结果（准备删除）'; en_US='Check every result (ready to delete)' }
 'Tip.Uncheck'   = @{ zh_TW='清掉清單中所有的勾選'; zh_CN='清除清单中所有勾选'; en_US='Clear all checkmarks' }
 'Tip.Delete'    = @{ zh_TW='刪除所有勾選項目。預設丟回收筒可還原；勾選上方「永久刪除」才會直接刪掉。'; zh_CN='删除所有勾选项。默认进回收站可还原；勾选上方「永久删除」才会直接删除。'; en_US='Delete all checked items. Goes to recycle bin by default; check Permanent delete to erase them directly.' }
 'Tip.Explorer'  = @{ zh_TW='在檔案總管中開啟目前選取的項目'; zh_CN='在文件资源管理器中打开当前选中项'; en_US='Open the selected item in Explorer' }
 'Tip.LogFolder' = @{ zh_TW='開啟操作日誌資料夾（logs\）'; zh_CN='打开操作日志文件夹（logs\）'; en_US='Open the log folder (logs\)' }
 'Tip.Export'    = @{ zh_TW='把目前搜尋結果存成 .txt / .csv / .json（格式依副檔名下拉自動套用）'; zh_CN='把当前搜索结果存成 .txt / .csv / .json（格式按扩展名下拉自动套用）'; en_US='Save results as .txt / .csv / .json (format follows the extension box)' }
 'Tip.Abort'     = @{ zh_TW='立即中止所有作業並結束程式（Ctrl+Q）。不再詢問；掃描或刪除會在目前這一步完成後停下，細節寫入日誌。'; zh_CN='立即中止所有作业并结束程序（Ctrl+Q）。不再询问；扫描或删除会在当前步骤完成后停下，细节写入日志。'; en_US='Abort all work and exit (Ctrl+Q). No further prompt; details are written to the log.' }
 'Tip.Refresh'   = @{ zh_TW='重新掃描磁碟機清單（插上 USB 後用這個）'; zh_CN='重新扫描磁盘列表（插上 USB 后用这个）'; en_US='Refresh the drive list (after plugging in USB)' }
 'Tip.ToggleDrives'= @{ zh_TW='切換全部磁碟機的勾選狀態'; zh_CN='切换全部磁盘的勾选状态'; en_US='Toggle the check state of all drives' }
 'Tip.UiLang'    = @{ zh_TW='切換本工具介面顯示的語言；不影響 ZW3D 軟體本身，也不影響搜尋結果。選擇會被記住。'; zh_CN='切换本工具界面显示的语言；不影响 ZW3D 软件本身，也不影响搜索结果。选择会被记住。'; en_US='Change the language of this tool interface only. It does not change ZW3D itself or the search results. The choice is remembered.' }
 'Tip.ScanLang'  = @{ zh_TW='選擇掃描時要辨識的 ZW3D 語言版本：影響常用關鍵字與在地化掃描關鍵詞；「全部語言」= 不限制。不影響介面文案。'; zh_CN='选择扫描时要识别的 ZW3D 语言版本：影响常用关键字与本地化扫描关键词；「全部语言」= 不限制。不影响界面文案。'; en_US='Choose the ZW3D language version used while scanning: affects keyword presets and localized scan words. All languages = no restriction. Does not change the UI text.' }
 'Drive.Available'= @{ zh_TW='({0} / {1} GB 可用)'; zh_CN='({0} / {1} GB 可用)'; en_US='({0} / {1} GB free)' }
 'Log.IconFail'  = @{ zh_TW='App 圖示建立失敗：{0}'; zh_CN='App 图标创建失败：{0}'; en_US='Failed to create app icon: {0}' }
 'Log.Drives'    = @{ zh_TW='磁碟機列舉完成，共 {0} 台'; zh_CN='磁盘枚举完成，共 {0} 台'; en_US='Drive enumeration done: {0}' }
 'Log.Start'     = @{ zh_TW='========== 檔案管理工具啟動 =========='; zh_CN='========== 文件管理工具启动 =========='; en_US='========== Folder Manager started ==========' }
 'Log.Close'     = @{ zh_TW='程式關閉'; zh_CN='程序关闭'; en_US='Application closed' }
 'Log.SearchThreadErr'= @{ zh_TW='搜尋執行緒異常：{0}'; zh_CN='搜索线程异常：{0}'; en_US='Search thread error: {0}' }
 'Log.SearchStart'= @{ zh_TW='開始搜尋：磁碟=[{0}] 關鍵字="{1}" 對象={2} 深度={3} ZW3D篩選={4} 副檔名={5} 掃描語言版本={6}'; zh_CN='开始搜索：磁盘=[{0}] 关键字="{1}" 对象={2} 深度={3} ZW3D筛选={4} 扩展名={5} 扫描语言版本={6}'; en_US='Search started: drives=[{0}] keyword="{1}" target={2} depth={3} ZW3D filter={4} ext={5} scanLang={6}' }
 'Log.SearchEnd' = @{ zh_TW='搜尋結束：關鍵字="{0}" 對象={1} 結果={2} 掃描={3} 無權限={4}'; zh_CN='搜索结束：关键字="{0}" 对象={1} 结果={2} 扫描={3} 无权限={4}'; en_US='Search finished: keyword="{0}" target={1} results={2} scanned={3} denied={4}' }
 'Log.SearchStartFail'= @{ zh_TW='無法啟動搜尋：{0}'; zh_CN='无法启动搜索：{0}'; en_US='Failed to start search: {0}' }
 'Log.UserStop'  = @{ zh_TW='使用者要求停止搜尋'; zh_CN='用户要求停止搜索'; en_US='User stopped the search' }
 'Log.Abort'     = @{ zh_TW='使用者觸發「立即停止」：中止所有作業並結束程式'; zh_CN='用户触发「立即停止」：中止所有作业并结束程序'; en_US='User triggered Abort: stopping all work and exiting' }
 'Log.UiLang'    = @{ zh_TW='介面語言切換為：{0}'; zh_CN='界面语言切换为：{0}'; en_US='UI language switched to: {0}' }
 'Log.ScanLang'  = @{ zh_TW='掃描語言版本切換為：{0}'; zh_CN='扫描语言版本切换为：{0}'; en_US='Scan language switched to: {0}' }
 'Log.Protected' = @{ zh_TW='拒絕刪除系統保護路徑：{0}'; zh_CN='拒绝删除系统保护路径：{0}'; en_US='Refused to delete protected system paths: {0}' }
 'Log.DeleteCancel'= @{ zh_TW='使用者取消刪除，共 {0} 項'; zh_CN='用户取消删除，共 {0} 项'; en_US='User cancelled deletion, {0} item(s)' }
 'Log.DeleteStart'= @{ zh_TW='===== 開始刪除（{0}），共 {1} 項（略過巢狀子項 {2} 個）====='; zh_CN='===== 开始删除（{0}），共 {1} 项（略过嵌套子项 {2} 个）====='; en_US='===== Delete started ({0}), {1} item(s), {2} nested skipped =====' }
 'Log.DeleteOk'  = @{ zh_TW='刪除成功（{0}）：{1}'; zh_CN='删除成功（{0}）：{1}'; en_US='Deleted ({0}): {1}' }
 'Log.DeleteFail'= @{ zh_TW='刪除失敗（{0}）：{1} → {2}'; zh_CN='删除失败（{0}）：{1} → {2}'; en_US='Delete failed ({0}): {1} -> {2}' }
 'Log.DeleteEnd' = @{ zh_TW='===== 刪除結束：{0} ====='; zh_CN='===== 删除结束：{0} ====='; en_US='===== Delete finished: {0} =====' }
 'Log.DeleteAbort'= @{ zh_TW='刪除被「立即停止」中斷：已處理 {0}/{1} 項（成功 {2}、失敗 {3}），其餘 {4} 項未動'; zh_CN='删除被「立即停止」中断：已处理 {0}/{1} 项（成功 {2}、失败 {3}），其余 {4} 项未动'; en_US='Delete aborted: {0}/{1} processed ({2} ok, {3} failed), {4} untouched' }
 'Log.ExportOk'  = @{ zh_TW='匯出清單成功（{0}）：{1}，共 {2} 筆'; zh_CN='导出清单成功（{0}）：{1}，共 {2} 条'; en_US='Export succeeded ({0}): {1}, {2} row(s)' }
 'Log.ExportFail'= @{ zh_TW='匯出清單失敗：{0} → {1}'; zh_CN='导出清单失败：{0} → {1}'; en_US='Export failed: {0} -> {1}' }
 'Log.OpenFail'  = @{ zh_TW='開啟失敗：{0}'; zh_CN='打开失败：{0}'; en_US='Open failed: {0}' }
 'Status.ReadyFull'= @{ zh_TW='就緒 · 勾選磁碟機後輸入關鍵字搜尋'; zh_CN='就绪 · 勾选磁盘后输入关键字搜索'; en_US='Ready · select a drive and enter a keyword to search' }
 'Status.Scanning'= @{ zh_TW='掃描中…'; zh_CN='扫描中…'; en_US='Scanning…' }
 'Status.ScanProgress'= @{ zh_TW='掃描中… 已檢查 {0} 項，命中 {1} 筆，略過 {2} 個無權限資料夾'; zh_CN='扫描中… 已检查 {0} 项，命中 {1} 条，略过 {2} 个无权限文件夹'; en_US='Scanning… {0} checked, {1} hits, {2} inaccessible folders skipped' }
 'Status.Found'  = @{ zh_TW='找到 {0} 筆（掃描 {1} 個項目，略過 {2} 個無權限資料夾）'; zh_CN='找到 {0} 条（扫描 {1} 个项目，略过 {2} 个无权限文件夹）'; en_US='{0} found ({1} scanned, {2} inaccessible folders skipped)' }
 'Status.Truncated'= @{ zh_TW=' · 已達結果上限 5000 筆，請縮小範圍'; zh_CN=' · 已达结果上限 5000 条，请缩小范围'; en_US=' · Result limit of 5000 reached, narrow the scope' }
 'Status.ExtHint'= @{ zh_TW=' · 提示：{0} 不在 ZW3D 篩選清單內，可取消勾選「僅顯示 ZW3D 相關」'; zh_CN=' · 提示：{0} 不在 ZW3D 筛选清单内，可取消勾选「仅显示 ZW3D 相关」'; en_US=' · Hint: {0} is not in the ZW3D filter list; uncheck ZW3D files only' }
 'Status.Checked'= @{ zh_TW='已勾選 {0} 項'; zh_CN='已勾选 {0} 项'; en_US='{0} item(s) checked' }
 'Status.Unchecked'= @{ zh_TW='已清除勾選'; zh_CN='已清除勾选'; en_US='Selection cleared' }
 'Status.DrivesRefreshed'= @{ zh_TW='磁碟機已重新整理'; zh_CN='磁盘已刷新'; en_US='Drive list refreshed' }
 'Status.Exported' = @{ zh_TW='已匯出 {0} 筆 → {1}'; zh_CN='已导出 {0} 条 → {1}'; en_US='Exported {0} row(s) -> {1}' }
 'Status.NoDelete'= @{ zh_TW='沒有可刪除的項目'; zh_CN='没有可删除的项目'; en_US='Nothing to delete' }
 'Status.Cancelled'= @{ zh_TW='已取消刪除'; zh_CN='已取消删除'; en_US='Deletion cancelled' }
 'Status.SearchFail'= @{ zh_TW='搜尋啟動失敗，請看日誌'; zh_CN='搜索启动失败，请看日志'; en_US='Failed to start search, see the log' }
 'Status.Aborting'= @{ zh_TW='立即停止：正在中止所有作業…'; zh_CN='立即停止：正在中止所有作业…'; en_US='Abort: stopping all work…' }
 'Status.Deleting'= @{ zh_TW='刪除中 {0}/{1}：{2}'; zh_CN='删除中 {0}/{1}：{2}'; en_US='Deleting {0}/{1}: {2}' }
 'Status.Summary'= @{ zh_TW='成功 {0} 項，失敗 {1} 項（{2}）'; zh_CN='成功 {0} 项，失败 {1} 项（{2}）'; en_US='{0} succeeded, {1} failed ({2})' }
 'Msg.InfoTitle' = @{ zh_TW='提示'; zh_CN='提示'; en_US='Information' }
 'Msg.ErrTitle'  = @{ zh_TW='錯誤'; zh_CN='错误'; en_US='Error' }
 'Msg.WarnTitle' = @{ zh_TW='警告'; zh_CN='警告'; en_US='Warning' }
 'Msg.NoDrive'   = @{ zh_TW='請至少勾選一個磁碟機。'; zh_CN='请至少勾选一个磁盘。'; en_US='Select at least one drive.' }
 'Msg.NoExport'  = @{ zh_TW='目前沒有可匯出的結果，請先搜尋。'; zh_CN='目前没有可导出的结果，请先搜索。'; en_US='No results to export. Run a search first.' }
 'Msg.NoSelection'= @{ zh_TW='請先勾選要刪除的項目（清單左側的核取方塊）。'; zh_CN='请先勾选要删除的项目（清单左侧的复选框）。'; en_US='Check the items you want to delete first.' }
 'Msg.ProtectedTitle'= @{ zh_TW='已跳過保護路徑'; zh_CN='已跳过保护路径'; en_US='Protected paths skipped' }
 'Msg.ProtectedBody'= @{ zh_TW="以下為系統保護路徑，已自動跳過：`r`n"; zh_CN="以下为系统保护路径，已自动跳过：`r`n"; en_US="The following protected system paths were skipped:`r`n" }
 'Msg.ExportOkTitle'= @{ zh_TW='匯出完成'; zh_CN='导出完成'; en_US='Export complete' }
 'Msg.ExportOk'  = @{ zh_TW="已匯出 {0} 筆至：`r`n{1}"; zh_CN="已导出 {0} 条至：`r`n{1}"; en_US="Exported {0} row(s) to:`r`n{1}" }
 'Msg.ExportFailBody'= @{ zh_TW="匯出失敗：`r`n"; zh_CN="导出失败：`r`n"; en_US="Export failed:`r`n" }
 'Msg.PartialTitle'= @{ zh_TW='部分項目刪除失敗'; zh_CN='部分项目删除失败'; en_US='Some items could not be deleted' }
 'Msg.ResultTitle'= @{ zh_TW='完成'; zh_CN='完成'; en_US='Done' }
 'Msg.ResultBody'= @{ zh_TW="刪除完成。`r`n`r`n{0}"; zh_CN="删除完成。`r`n`r`n{0}"; en_US="Deletion complete.`r`n`r`n{0}" }
 'Msg.FailListBody'= @{ zh_TW="刪除結果：{0}`r`n`r`n失敗項目：`r`n"; zh_CN="删除结果：{0}`r`n`r`n失败项目：`r`n"; en_US="Delete result: {0}`r`n`r`nFailed items:`r`n" }
 'Msg.RecycleFail'= @{ zh_TW='回收筒元件不可用，請改勾「永久刪除」後重試'; zh_CN='回收站组件不可用，请改勾「永久删除」后重试'; en_US='Recycle bin component unavailable; check Permanent delete and retry' }
 'Msg.ErrPermission'= @{ zh_TW='權限不足（UnauthorizedAccess）'; zh_CN='权限不足（UnauthorizedAccess）'; en_US='Permission denied (UnauthorizedAccess)' }
 'Msg.ErrSecurity'= @{ zh_TW='安全性限制（SecurityException）'; zh_CN='安全性限制（SecurityException）'; en_US='Security restriction (SecurityException)' }
 'Msg.ErrLocked' = @{ zh_TW='檔案／資料夾被其他程式佔用：'; zh_CN='文件／文件夹被其他程序占用：'; en_US='File / folder in use by another process: ' }
 'Msg.ErrLongPath'= @{ zh_TW='路徑過長或無效：'; zh_CN='路径过长或无效：'; en_US='Path too long or invalid: ' }
 'Msg.ErrIO'     = @{ zh_TW='IO 錯誤：'; zh_CN='IO 错误：'; en_US='I/O error: ' }
 'Msg.ErrUnexpected'= @{ zh_TW='未預期錯誤：{0}'; zh_CN='未预期错误：{0}'; en_US='Unexpected error: {0}' }
 'Dlg.Title'     = @{ zh_TW='確認刪除'; zh_CN='确认删除'; en_US='Confirm deletion' }
 'Dlg.Heading'   = @{ zh_TW='即將刪除 {0} 個項目'; zh_CN='即将删除 {0} 个项目'; en_US='About to delete {0} item(s)' }
 'Dlg.Perm'      = @{ zh_TW='模式：永久刪除。檔案不會進入回收筒，刪除後無法復原。'; zh_CN='模式：永久删除。文件不会进入回收站，删除后无法恢复。'; en_US='Mode: permanent delete. Files bypass the recycle bin and cannot be restored.' }
 'Dlg.Recycle'   = @{ zh_TW='模式：移至資源回收筒。可從回收筒還原。'; zh_CN='模式：移至回收站。可从回收站还原。'; en_US='Mode: move to recycle bin. Items can be restored.' }
 'Dlg.Confirm'   = @{ zh_TW='我確認要刪除上述項目'; zh_CN='我确认要删除上述项目'; en_US='I confirm I want to delete the items above' }
 'Dlg.Cancel'    = @{ zh_TW='取消'; zh_CN='取消'; en_US='Cancel' }
 'Dlg.Ok'        = @{ zh_TW='確定刪除'; zh_CN='确定删除'; en_US='Delete' }
 'Export.Title'  = @{ zh_TW='匯出搜尋結果清單'; zh_CN='导出搜索结果清单'; en_US='Export search results' }
 'Export.FilterTxt'= @{ zh_TW='文字檔 (*.txt)|*.txt'; zh_CN='文本文件 (*.txt)|*.txt'; en_US='Text file (*.txt)|*.txt' }
 'Export.FilterJson'= @{ zh_TW='JSON (*.json)|*.json'; zh_CN='JSON (*.json)|*.json'; en_US='JSON (*.json)|*.json' }
 'Export.FilterCsv'= @{ zh_TW='CSV (*.csv)|*.csv'; zh_CN='CSV (*.csv)|*.csv'; en_US='CSV (*.csv)|*.csv' }
 'Export.FileName'= @{ zh_TW='ZW3D清單_'; zh_CN='ZW3D清单_'; en_US='ZW3D_list_' }
 'Mode.Permanent'= @{ zh_TW='永久刪除'; zh_CN='永久删除'; en_US='permanent delete' }
 'Mode.Recycle'  = @{ zh_TW='移至回收筒'; zh_CN='移至回收站'; en_US='moved to recycle bin' }
 'ZwMode.Off'    = @{ zh_TW='關閉'; zh_CN='关闭'; en_US='off' }
 'ZwMode.All'    = @{ zh_TW='全部 ZW3D 檔案'; zh_CN='全部 ZW3D 文件'; en_US='all ZW3D files' }
 'ZwMode.Junk'   = @{ zh_TW='僅 ZW3D 暫存垃圾檔'; zh_CN='仅 ZW3D 暂存垃圾文件'; en_US='ZW3D temp junk only' }
 'ZwMode.Model'  = @{ zh_TW='僅 ZW3D 模型檔'; zh_CN='仅 ZW3D 模型文件'; en_US='ZW3D model files only' }
 'Btn.UserFolder' = @{ zh_TW='使用者目錄'; zh_CN='用户目录'; en_US='User folder'; de_DE='Benutzerordner'; fr_FR='Dossier utilisateur'; es_ES='Carpeta de usuario'; it_IT='Cartella utente'; pt_PT='Pasta do utilizador'; ja_JP='ユーザーフォルダー'; ko_KR='사용자 폴더'; ru_RU='Папка пользователя'; pl_PL='Folder użytkownika'; cs_CZ='Složka uživatele'; hu_HU='Felhasználói mappa'; tr_TR='Kullanıcı klasörü' }
 'Tip.UserFolder' = @{ zh_TW='清理 ZW3D 使用者目錄，提供與官方 User Folder Manager 相同的四個動作'; zh_CN='清理 ZW3D 用户目录，提供与官方 User Folder Manager 相同的四个动作'; en_US='Clean the ZW3D user folder with the same four actions as the official User Folder Manager' }
 'Dlg.UfTitle'    = @{ zh_TW='ZW3D 使用者目錄清理'; zh_CN='ZW3D 用户目录清理'; en_US='ZW3D User Folder Manager'; de_DE='ZW3D Benutzerordner'; fr_FR='Dossier utilisateur ZW3D'; es_ES='Carpeta de usuario ZW3D'; ja_JP='ZW3D ユーザーフォルダー'; ko_KR='ZW3D 사용자 폴더' }
 'Dlg.UfVersion'  = @{ zh_TW='選擇要刪除的版本：'; zh_CN='选择要删除的版本：'; en_US='Version to clean:'; de_DE='Zu bereinigende Version:'; fr_FR='Version a nettoyer :'; es_ES='Version a limpiar:'; it_IT='Versione da pulire:'; pt_PT='Versao a limpar:'; ja_JP='削除するバージョン：'; ko_KR='정리할 버전:'; ru_RU='Версия для очистки:'; pl_PL='Wersja do czyszczenia:'; cs_CZ='Verze k cisteni:'; hu_HU='Tisztitando verzio:'; tr_TR='Temizlenecek surum:' }
 'Dlg.UfLocation' = @{ zh_TW='位置：'; zh_CN='位置：'; en_US='Location:'; de_DE='Ort:'; fr_FR='Emplacement :'; es_ES='Ubicacion:'; it_IT='Posizione:'; pt_PT='Localizacao:'; ja_JP='場所：'; ko_KR='위치:'; ru_RU='Расположение:'; pl_PL='Lokalizacja:'; cs_CZ='Umisteni:'; hu_HU='Hely:'; tr_TR='Konum:' }
 'Dlg.UfHint'     = @{ zh_TW='已自動偵測本機的 ZW3D 版本使用者目錄；切換版本會自動帶出對應路徑，無需手動輸入。'; zh_CN='已自动检测本机的 ZW3D 版本用户目录；切换版本会自动带出对应路径，无需手动输入。'; en_US='Installed ZW3D versions are detected automatically. Switching the version updates the paths below - no typing needed.'; de_DE='Installierte ZW3D-Versionen werden automatisch erkannt; beim Wechseln aktualisieren sich die Pfade.'; fr_FR='Versions ZW3D detectees automatiquement; les chemins se mettent a jour lors du changement.'; es_ES='Versiones ZW3D detectadas automaticamente; las rutas se actualizan al cambiar de version.'; it_IT='Versioni ZW3D rilevate automaticamente; i percorsi si aggiornano cambiando versione.'; pt_PT='Versoes ZW3D detetadas automaticamente; os caminhos atualizam ao mudar de versao.'; ja_JP='インストール済みの ZW3D バージョンを自動検出します。切り替えるとパスも自動更新されます。'; ko_KR='설치된 ZW3D 버전을 자동으로 감지합니다. 버전을 바꾸면 경로도 자동으로 갱신됩니다.'; ru_RU='Установленные версии ZW3D определяются автоматически; при переключении пути обновляются.'; pl_PL='Zainstalowane wersje ZW3D sa wykrywane automatycznie; sciezki aktualizuja sie po zmianie wersji.'; cs_CZ='Nainstalovane verze ZW3D jsou detekovany automaticky; cesty se aktualizuji pri zmene verze.'; hu_HU='A telepitett ZW3D verziok automatikusan felismeresre kerulnek; az utvonalak valtaskor frissulnek.'; tr_TR='Kurulu ZW3D surumleri otomatik algilanir; surum degisince yollar guncellenir.' }
 'Dlg.UfOpen'     = @{ zh_TW='開啟'; zh_CN='打开'; en_US='Open'; de_DE='Offnen'; fr_FR='Ouvrir'; es_ES='Abrir'; it_IT='Apri'; pt_PT='Abrir'; ja_JP='開く'; ko_KR='열기'; ru_RU='Открыть'; pl_PL='Otworz'; cs_CZ='Otevrit'; hu_HU='Megnyitas'; tr_TR='Ac' }
 'Dlg.UfClose'    = @{ zh_TW='關閉'; zh_CN='关闭'; en_US='Close'; de_DE='Schliessen'; fr_FR='Fermer'; es_ES='Cerrar'; it_IT='Chiudi'; pt_PT='Fechar'; ja_JP='閉じる'; ko_KR='닫기'; ru_RU='Закрыть'; pl_PL='Zamknij'; cs_CZ='Zavrit'; hu_HU='Bezaras'; tr_TR='Kapat' }
 'Dlg.UfNote'     = @{ zh_TW='注意：清理會影響執行中的 ZW3D，請先關閉。刪除前會再次確認。'; zh_CN='注意：清理会影响运行中的 ZW3D，请先关闭。删除前会再次确认。'; en_US='Note: clearing will affect the running ZW3D; close it first. A confirmation is shown before deleting.' }
 'Uf.StatLine'    = @{ zh_TW='共 {0} 個檔案 · {1} 個資料夾 · {2}'; zh_CN='共 {0} 个文件 · {1} 个文件夹 · {2}'; en_US='{0} file(s) / {1} folder(s) / {2}' }
 'Uf.NoTarget'    = @{ zh_TW='（目錄不存在，無需清理）'; zh_CN='（目录不存在，无需清理）'; en_US='(folder not found, nothing to clean)' }
 'Uf.ClearOutput' = @{ zh_TW='清除自動輸出檔案'; zh_CN='清除自动输出文件'; en_US='Clear auto output files' }
 'Uf.ClearCustom' = @{ zh_TW='清除自訂組態'; zh_CN='清除自定义配置'; en_US='Clear custom configurations' }
 'Uf.ClearTemp'   = @{ zh_TW='清除暫存檔'; zh_CN='清除临时文件'; en_US='Clear temp files' }
 'Uf.ClearAll'    = @{ zh_TW='清除使用者目錄'; zh_CN='清除用户目录'; en_US='Clear user folder' }
 'Uf.DescOutput'  = @{ zh_TW='刪除 output\（日誌、封存、轉換輸出等）'; zh_CN='删除 output\（日志、归档、转换输出等）'; en_US='Delete output\ (logs, archives, converted files)' }
 'Uf.DescCustom'  = @{ zh_TW='刪除 custom\（介面組態、範本、授權等，會重置個人設定）'; zh_CN='删除 custom\（界面配置、模板、授权等，会重置个人设置）'; en_US='Delete custom\ (UI config, templates, licenses - resets personal settings)' }
 'Uf.DescTemp'    = @{ zh_TW='刪除 temp\（工作階段與中繼暫存檔）'; zh_CN='删除 temp\（会话与中间临时文件）'; en_US='Delete temp\ (session and intermediate files)' }
 'Uf.DescAll'     = @{ zh_TW='刪除整個版本目錄（含上述全部與其他檔案）'; zh_CN='删除整个版本目录（含上述全部与其他文件）'; en_US='Delete the whole instance folder (everything above and more)' }
 'Uf.Confirm'     = @{ zh_TW="即將執行：{0}`r`n版本：{5}`r`n目錄：{1}`r`n共 {2} 個檔案、{3} 個子資料夾。`r`n模式：{4}"; zh_CN="即将执行：{0}`r`n版本：{5}`r`n目录：{1}`r`n共 {2} 个文件、{3} 个子文件夹。`r`n模式：{4}"; en_US="Action: {0}`r`nVersion: {5}`r`nFolder: {1}`r`n{2} file(s), {3} subfolder(s).`r`nMode: {4}" }
 'Uf.WarnCustom'  = @{ zh_TW="`r`n`r`n警告：這會刪除個人組態，ZW3D 下次啟動會重建。"; zh_CN="`r`n`r`n警告：这会删除个人配置，ZW3D 下次启动会重建。"; en_US="`r`n`r`nWarning: personal configurations will be removed; ZW3D rebuilds them on next start." }
 'Uf.WarnAll'     = @{ zh_TW="`r`n`r`n警告：這會清空整個使用者目錄，請先備份需要的設定。"; zh_CN="`r`n`r`n警告：这会清空整个用户目录，请先备份需要的设置。"; en_US="`r`n`r`nWarning: the whole user folder will be emptied. Back up anything you need first." }
 'Uf.ZwRunning'   = @{ zh_TW='偵測到 ZW3D 正在執行，檔案可能被鎖定。建議先關閉 ZW3D。仍要繼續？'; zh_CN='检测到 ZW3D 正在运行，文件可能被锁定。建议先关闭 ZW3D。仍要继续？'; en_US='ZW3D is running; files may be locked. Close ZW3D first. Continue anyway?' }
 'Uf.NoFolder'    = @{ zh_TW='找不到 ZW3D 使用者目錄：{0}'; zh_CN='找不到 ZW3D 用户目录：{0}'; en_US='ZW3D user folder not found: {0}' }
 'Uf.Missing'     = @{ zh_TW='目錄不存在，無需清理：{0}'; zh_CN='目录不存在，无需清理：{0}'; en_US='Folder does not exist, nothing to clean: {0}' }
 'Uf.Ok'          = @{ zh_TW='清理完成：{0}'; zh_CN='清理完成：{0}'; en_US='Cleaning done: {0}' }
 'Uf.Fail'        = @{ zh_TW='清理失敗：{0}'; zh_CN='清理失败：{0}'; en_US='Cleaning failed: {0}' }
 'Uf.Cancelled'   = @{ zh_TW='已取消'; zh_CN='已取消'; en_US='Cancelled' }
 'Uf.ConfirmTitle'= @{ zh_TW='確認清理'; zh_CN='确认清理'; en_US='Confirm cleaning' }
 'Log.UserClean'  = @{ zh_TW='使用者目錄清理：動作={0} 目錄={1} 結果={2}'; zh_CN='用户目录清理：动作={0} 目录={1} 结果={2}'; en_US='User folder cleaning: action={0} folder={1} result={2}' }
 'Ext.Unlimited' = @{ zh_TW='不限'; zh_CN='不限'; en_US='all' }
}

# 補充詞條：依語言分組（較易維護）。所有畫面文字、視窗／對話框標題、提示、狀態、訊息都在此補齊 15 種語言；
# 日誌（Log.*）刻意維持英文回退，方便跨語言比對與回報問題。'¶' 代表換行（載入時轉成 CRLF）。
$script:StrMore = @{
 'zh_TW' = @{ 'Lang.All' = '全部語言' }
 'zh_CN' = @{ 'Lang.All' = '全部语言' }
 'en_US' = @{ 'Lang.All' = 'All languages' }
 'de_DE' = @{
  'Lang.All' = 'Alle Sprachen'
  'Hint' = 'Tipp: Stichwort leer = alles auflisten (max. 5000 Treffer). Enter startet die Suche.'
  'Zw3dHint' = 'Kriterien: Erweiterungsliste / Name enthält ZW3D·TopologicRep·Proxy / Pfad enthält ZW3D'
  'ZwScope.All' = 'Alle ZW3D-Dateien (Modelle + Temp)'
  'ZwScope.Junk' = 'Nur ZW3D-Temp-Müll (clear ALL.bat)'
  'ZwScope.Model' = 'Nur ZW3D-Modelle (.z3/.z3prt/.z3asm/.z3drw)'
  'CreditTip' = 'Entwickler'
  'Tip.Search' = 'Suche starten (oder Enter im Stichwortfeld drücken)'
  'Tip.Stop' = 'Laufende Suche anhalten'
  'Tip.CheckAll' = 'Alle Ergebnisse anhaken (zum Löschen vormerken)'
  'Tip.Uncheck' = 'Alle Häkchen entfernen'
  'Tip.Delete' = 'Alle angehakten Elemente löschen. Standard: Papierkorb (wiederherstellbar); mit „Endgültig löschen“ werden sie direkt entfernt.'
  'Tip.Explorer' = 'Ausgewähltes Element im Explorer öffnen'
  'Tip.LogFolder' = 'Protokollordner öffnen (logs\)'
  'Tip.Export' = 'Ergebnisse als .txt / .csv / .json speichern (Format folgt dem Erweiterungsfeld)'
  'Tip.Abort' = 'Alle Vorgänge abbrechen und beenden (Strg+Q). Keine weitere Rückfrage; Details stehen im Protokoll.'
  'Tip.Refresh' = 'Laufwerksliste neu einlesen (nach dem Anschließen eines USB-Laufwerks)'
  'Tip.ToggleDrives' = 'Auswahl aller Laufwerke umschalten'
  'Tip.UiLang' = 'Ändert nur die Sprache dieser Oberfläche – nicht ZW3D selbst und nicht die Suchergebnisse. Die Auswahl wird gespeichert.'
  'Tip.ScanLang' = 'ZW3D-Sprachversion für die Suche: beeinflusst Stichwortvorlagen und lokalisierte Suchbegriffe. „Alle Sprachen“ = keine Einschränkung. Ändert nicht die Oberfläche.'
  'Tip.UserFolder' = 'ZW3D-Benutzerordner bereinigen – mit denselben vier Aktionen wie der offizielle User Folder Manager'
  'Drive.Available' = '({0} / {1} GB frei)'
  'Status.ReadyFull' = 'Bereit · Laufwerk wählen und Stichwort eingeben'
  'Status.Scanning' = 'Suche läuft…'
  'Status.ScanProgress' = 'Suche läuft… {0} geprüft, {1} Treffer, {2} gesperrte Ordner übersprungen'
  'Status.Found' = '{0} gefunden ({1} geprüft, {2} gesperrte Ordner übersprungen)'
  'Status.Truncated' = ' · Limit von 5000 Treffern erreicht, bitte Suche eingrenzen'
  'Status.ExtHint' = ' · Hinweis: {0} ist nicht in der ZW3D-Filterliste; „Nur ZW3D-Dateien“ abwählen'
  'Status.Checked' = '{0} Element(e) angehakt'
  'Status.Unchecked' = 'Auswahl aufgehoben'
  'Status.DrivesRefreshed' = 'Laufwerksliste aktualisiert'
  'Status.Exported' = '{0} Zeile(n) exportiert -> {1}'
  'Status.NoDelete' = 'Nichts zu löschen'
  'Status.Cancelled' = 'Löschen abgebrochen'
  'Status.SearchFail' = 'Suche konnte nicht gestartet werden, siehe Protokoll'
  'Status.Aborting' = 'Abbruch: alle Vorgänge werden beendet…'
  'Status.Deleting' = 'Lösche {0}/{1}: {2}'
  'Status.Summary' = '{0} erfolgreich, {1} fehlgeschlagen ({2})'
  'Msg.InfoTitle' = 'Hinweis'
  'Msg.ErrTitle' = 'Fehler'
  'Msg.WarnTitle' = 'Warnung'
  'Msg.NoDrive' = 'Bitte mindestens ein Laufwerk auswählen.'
  'Msg.NoExport' = 'Keine Ergebnisse zum Exportieren. Bitte zuerst suchen.'
  'Msg.NoSelection' = 'Bitte zuerst die zu löschenden Elemente anhaken.'
  'Msg.ProtectedTitle' = 'Geschützte Pfade übersprungen'
  'Msg.ProtectedBody' = 'Folgende geschützte Systempfade wurden übersprungen:¶'
  'Msg.ExportOkTitle' = 'Export abgeschlossen'
  'Msg.ExportOk' = '{0} Zeile(n) exportiert nach:¶{1}'
  'Msg.ExportFailBody' = 'Export fehlgeschlagen:¶'
  'Msg.PartialTitle' = 'Einige Elemente konnten nicht gelöscht werden'
  'Msg.ResultTitle' = 'Fertig'
  'Msg.ResultBody' = 'Löschen abgeschlossen.¶¶{0}'
  'Msg.FailListBody' = 'Ergebnis: {0}¶¶Fehlgeschlagene Elemente:¶'
  'Msg.RecycleFail' = 'Papierkorb-Komponente nicht verfügbar; „Endgültig löschen“ aktivieren und erneut versuchen'
  'Msg.ErrPermission' = 'Zugriff verweigert (UnauthorizedAccess)'
  'Msg.ErrSecurity' = 'Sicherheitsbeschränkung (SecurityException)'
  'Msg.ErrLocked' = 'Datei / Ordner wird von einem anderen Prozess verwendet: '
  'Msg.ErrLongPath' = 'Pfad zu lang oder ungültig: '
  'Msg.ErrIO' = 'E/A-Fehler: '
  'Msg.ErrUnexpected' = 'Unerwarteter Fehler: {0}'
  'Dlg.Title' = 'Löschen bestätigen'
  'Dlg.Heading' = '{0} Element(e) werden gelöscht'
  'Dlg.Perm' = 'Modus: endgültig löschen. Dateien umgehen den Papierkorb und können nicht wiederhergestellt werden.'
  'Dlg.Recycle' = 'Modus: in den Papierkorb verschieben. Wiederherstellung möglich.'
  'Dlg.Confirm' = 'Ich bestätige das Löschen der obigen Elemente'
  'Dlg.Cancel' = 'Abbrechen'
  'Dlg.Ok' = 'Löschen'
  'Export.Title' = 'Suchergebnisse exportieren'
  'Export.FilterTxt' = 'Textdatei (*.txt)|*.txt'
  'Export.FileName' = 'ZW3D_Liste_'
  'Mode.Permanent' = 'endgültig löschen'
  'Mode.Recycle' = 'in den Papierkorb'
  'ZwMode.Off' = 'aus'
  'ZwMode.All' = 'alle ZW3D-Dateien'
  'ZwMode.Junk' = 'nur ZW3D-Temp-Müll'
  'ZwMode.Model' = 'nur ZW3D-Modelle'
  'Dlg.UfTitle' = 'ZW3D-Benutzerordner bereinigen'
  'Dlg.UfVersion' = 'Zu bereinigende Version:'
  'Dlg.UfLocation' = 'Ort:'
  'Dlg.UfHint' = 'Installierte ZW3D-Versionen werden automatisch erkannt; beim Wechsel der Version werden die Pfade unten aktualisiert.'
  'Dlg.UfOpen' = 'Öffnen'
  'Dlg.UfClose' = 'Schließen'
  'Dlg.UfNote' = 'Hinweis: Das Bereinigen betrifft ein laufendes ZW3D – bitte vorher schließen. Vor dem Löschen erfolgt eine Rückfrage.'
  'Uf.StatLine' = '{0} Datei(en) / {1} Ordner / {2}'
  'Uf.NoTarget' = '(Ordner nicht vorhanden, nichts zu bereinigen)'
  'Uf.ClearOutput' = 'Ausgabedateien löschen'
  'Uf.ClearCustom' = 'Eigene Konfiguration löschen'
  'Uf.ClearTemp' = 'Temp-Dateien löschen'
  'Uf.ClearAll' = 'Benutzerordner leeren'
  'Uf.DescOutput' = 'Löscht output\ (Protokolle, Archive, konvertierte Dateien)'
  'Uf.DescCustom' = 'Löscht custom\ (UI-Konfiguration, Vorlagen, Lizenzen – setzt persönliche Einstellungen zurück)'
  'Uf.DescTemp' = 'Löscht temp\ (Sitzungs- und Zwischendateien)'
  'Uf.DescAll' = 'Löscht den gesamten Versionsordner (alles oben und mehr)'
  'Uf.Confirm' = 'Aktion: {0}¶Version: {5}¶Ordner: {1}¶{2} Datei(en), {3} Unterordner.¶Modus: {4}'
  'Uf.WarnCustom' = '¶¶Warnung: Persönliche Konfigurationen werden entfernt; ZW3D erstellt sie beim nächsten Start neu.'
  'Uf.WarnAll' = '¶¶Warnung: Der gesamte Benutzerordner wird geleert. Benötigte Einstellungen vorher sichern.'
  'Uf.ZwRunning' = 'ZW3D läuft; Dateien können gesperrt sein. Bitte ZW3D zuerst schließen. Trotzdem fortfahren?'
  'Uf.NoFolder' = 'ZW3D-Benutzerordner nicht gefunden: {0}'
  'Uf.Missing' = 'Ordner nicht vorhanden, nichts zu bereinigen: {0}'
  'Uf.Ok' = 'Bereinigung abgeschlossen: {0}'
  'Uf.Fail' = 'Bereinigung fehlgeschlagen: {0}'
  'Uf.Cancelled' = 'Abgebrochen'
  'Uf.ConfirmTitle' = 'Bereinigung bestätigen'
 }
 'fr_FR' = @{
  'Lang.All' = 'Toutes les langues'
  'Hint' = 'Astuce : mot-clé vide = tout lister (max. 5000 résultats). Entrée lance la recherche.'
  'Zw3dHint' = 'Critères : liste d''extensions / nom contient ZW3D·TopologicRep·Proxy / chemin contient ZW3D'
  'ZwScope.All' = 'Tous les fichiers ZW3D (modèles + temp)'
  'ZwScope.Junk' = 'Fichiers temporaires ZW3D (clear ALL.bat)'
  'ZwScope.Model' = 'Modèles ZW3D uniquement (.z3/.z3prt/.z3asm/.z3drw)'
  'CreditTip' = 'Développeur'
  'Tip.Search' = 'Lancer la recherche (ou Entrée dans le champ mot-clé)'
  'Tip.Stop' = 'Arrêter l''analyse en cours'
  'Tip.CheckAll' = 'Cocher tous les résultats (prêts à supprimer)'
  'Tip.Uncheck' = 'Décocher tout'
  'Tip.Delete' = 'Supprimer les éléments cochés. Par défaut vers la corbeille (restaurable) ; cochez « Suppression définitive » pour les effacer directement.'
  'Tip.Explorer' = 'Ouvrir l''élément sélectionné dans l''Explorateur'
  'Tip.LogFolder' = 'Ouvrir le dossier des journaux (logs\)'
  'Tip.Export' = 'Enregistrer les résultats en .txt / .csv / .json (format selon la liste d''extensions)'
  'Tip.Abort' = 'Interrompre toutes les opérations et quitter (Ctrl+Q). Sans confirmation ; détails dans le journal.'
  'Tip.Refresh' = 'Actualiser la liste des disques (après branchement d''une clé USB)'
  'Tip.ToggleDrives' = 'Inverser la sélection de tous les disques'
  'Tip.UiLang' = 'Change uniquement la langue de cette interface, pas ZW3D ni les résultats. Le choix est mémorisé.'
  'Tip.ScanLang' = 'Version linguistique de ZW3D pour l''analyse : agit sur les mots-clés prédéfinis et les termes localisés. « Toutes les langues » = sans restriction. Ne change pas l''interface.'
  'Tip.UserFolder' = 'Nettoyer le dossier utilisateur ZW3D avec les quatre actions du User Folder Manager officiel'
  'Drive.Available' = '({0} / {1} Go libres)'
  'Status.ReadyFull' = 'Prêt · choisissez un disque et saisissez un mot-clé'
  'Status.Scanning' = 'Analyse…'
  'Status.ScanProgress' = 'Analyse… {0} vérifiés, {1} trouvés, {2} dossiers inaccessibles ignorés'
  'Status.Found' = '{0} trouvé(s) ({1} analysés, {2} dossiers inaccessibles ignorés)'
  'Status.Truncated' = ' · Limite de 5000 résultats atteinte, affinez la recherche'
  'Status.ExtHint' = ' · Astuce : {0} n''est pas dans le filtre ZW3D ; décochez « Fichiers ZW3D uniquement »'
  'Status.Checked' = '{0} élément(s) coché(s)'
  'Status.Unchecked' = 'Sélection effacée'
  'Status.DrivesRefreshed' = 'Liste des disques actualisée'
  'Status.Exported' = '{0} ligne(s) exportée(s) -> {1}'
  'Status.NoDelete' = 'Rien à supprimer'
  'Status.Cancelled' = 'Suppression annulée'
  'Status.SearchFail' = 'Impossible de lancer la recherche, voir le journal'
  'Status.Aborting' = 'Arrêt : interruption de toutes les opérations…'
  'Status.Deleting' = 'Suppression {0}/{1} : {2}'
  'Status.Summary' = '{0} réussi(s), {1} échec(s) ({2})'
  'Msg.InfoTitle' = 'Information'
  'Msg.ErrTitle' = 'Erreur'
  'Msg.WarnTitle' = 'Avertissement'
  'Msg.NoDrive' = 'Sélectionnez au moins un disque.'
  'Msg.NoExport' = 'Aucun résultat à exporter. Lancez d''abord une recherche.'
  'Msg.NoSelection' = 'Cochez d''abord les éléments à supprimer.'
  'Msg.ProtectedTitle' = 'Chemins protégés ignorés'
  'Msg.ProtectedBody' = 'Les chemins système protégés suivants ont été ignorés :¶'
  'Msg.ExportOkTitle' = 'Export terminé'
  'Msg.ExportOk' = '{0} ligne(s) exportée(s) vers :¶{1}'
  'Msg.ExportFailBody' = 'Échec de l''export :¶'
  'Msg.PartialTitle' = 'Certains éléments n''ont pas pu être supprimés'
  'Msg.ResultTitle' = 'Terminé'
  'Msg.ResultBody' = 'Suppression terminée.¶¶{0}'
  'Msg.FailListBody' = 'Résultat : {0}¶¶Éléments en échec :¶'
  'Msg.RecycleFail' = 'Composant Corbeille indisponible ; cochez « Suppression définitive » et réessayez'
  'Msg.ErrPermission' = 'Accès refusé (UnauthorizedAccess)'
  'Msg.ErrSecurity' = 'Restriction de sécurité (SecurityException)'
  'Msg.ErrLocked' = 'Fichier / dossier utilisé par un autre processus : '
  'Msg.ErrLongPath' = 'Chemin trop long ou invalide : '
  'Msg.ErrIO' = 'Erreur d''E/S : '
  'Msg.ErrUnexpected' = 'Erreur inattendue : {0}'
  'Dlg.Title' = 'Confirmer la suppression'
  'Dlg.Heading' = '{0} élément(s) vont être supprimés'
  'Dlg.Perm' = 'Mode : suppression définitive. Les fichiers ne passent pas par la corbeille et ne peuvent pas être restaurés.'
  'Dlg.Recycle' = 'Mode : déplacer vers la corbeille. Les éléments peuvent être restaurés.'
  'Dlg.Confirm' = 'Je confirme la suppression des éléments ci-dessus'
  'Dlg.Cancel' = 'Annuler'
  'Dlg.Ok' = 'Supprimer'
  'Export.Title' = 'Exporter les résultats'
  'Export.FilterTxt' = 'Fichier texte (*.txt)|*.txt'
  'Export.FileName' = 'Liste_ZW3D_'
  'Mode.Permanent' = 'suppression définitive'
  'Mode.Recycle' = 'vers la corbeille'
  'ZwMode.Off' = 'désactivé'
  'ZwMode.All' = 'tous les fichiers ZW3D'
  'ZwMode.Junk' = 'temporaires ZW3D uniquement'
  'ZwMode.Model' = 'modèles ZW3D uniquement'
  'Dlg.UfTitle' = 'Nettoyage du dossier utilisateur ZW3D'
  'Dlg.UfVersion' = 'Version à nettoyer :'
  'Dlg.UfLocation' = 'Emplacement :'
  'Dlg.UfHint' = 'Les versions de ZW3D installées sont détectées automatiquement ; changer de version met à jour les chemins ci-dessous.'
  'Dlg.UfOpen' = 'Ouvrir'
  'Dlg.UfClose' = 'Fermer'
  'Dlg.UfNote' = 'Remarque : le nettoyage affecte ZW3D en cours d''exécution ; fermez-le d''abord. Une confirmation est demandée avant suppression.'
  'Uf.StatLine' = '{0} fichier(s) / {1} dossier(s) / {2}'
  'Uf.NoTarget' = '(dossier introuvable, rien à nettoyer)'
  'Uf.ClearOutput' = 'Effacer les sorties'
  'Uf.ClearCustom' = 'Effacer la configuration'
  'Uf.ClearTemp' = 'Effacer les temporaires'
  'Uf.ClearAll' = 'Vider le dossier utilisateur'
  'Uf.DescOutput' = 'Supprime output\ (journaux, archives, fichiers convertis)'
  'Uf.DescCustom' = 'Supprime custom\ (config. interface, modèles, licences – réinitialise vos paramètres)'
  'Uf.DescTemp' = 'Supprime temp\ (fichiers de session et intermédiaires)'
  'Uf.DescAll' = 'Supprime tout le dossier de la version (tout ce qui précède et plus)'
  'Uf.Confirm' = 'Action : {0}¶Version : {5}¶Dossier : {1}¶{2} fichier(s), {3} sous-dossier(s).¶Mode : {4}'
  'Uf.WarnCustom' = '¶¶Attention : vos configurations personnelles seront supprimées ; ZW3D les recrée au prochain démarrage.'
  'Uf.WarnAll' = '¶¶Attention : tout le dossier utilisateur sera vidé. Sauvegardez d''abord ce dont vous avez besoin.'
  'Uf.ZwRunning' = 'ZW3D est en cours d''exécution ; des fichiers peuvent être verrouillés. Fermez d''abord ZW3D. Continuer quand même ?'
  'Uf.NoFolder' = 'Dossier utilisateur ZW3D introuvable : {0}'
  'Uf.Missing' = 'Le dossier n''existe pas, rien à nettoyer : {0}'
  'Uf.Ok' = 'Nettoyage terminé : {0}'
  'Uf.Fail' = 'Échec du nettoyage : {0}'
  'Uf.Cancelled' = 'Annulé'
  'Uf.ConfirmTitle' = 'Confirmer le nettoyage'
 }
 'es_ES' = @{
  'Lang.All' = 'Todos los idiomas'
  'Hint' = 'Consejo: palabra clave vacía = listar todo (máx. 5000 resultados). Intro para buscar.'
  'Zw3dHint' = 'Criterio: lista de extensiones / nombre con ZW3D·TopologicRep·Proxy / ruta con ZW3D'
  'ZwScope.All' = 'Todos los archivos ZW3D (modelos + temp.)'
  'ZwScope.Junk' = 'Solo temporales ZW3D (clear ALL.bat)'
  'ZwScope.Model' = 'Solo modelos ZW3D (.z3/.z3prt/.z3asm/.z3drw)'
  'CreditTip' = 'Desarrollador'
  'Tip.Search' = 'Iniciar búsqueda (o pulse Intro en el cuadro de palabra clave)'
  'Tip.Stop' = 'Detener el análisis en curso'
  'Tip.CheckAll' = 'Marcar todos los resultados (listos para eliminar)'
  'Tip.Uncheck' = 'Quitar todas las marcas'
  'Tip.Delete' = 'Eliminar los elementos marcados. Por defecto van a la papelera (recuperables); marque «Eliminar permanentemente» para borrarlos directamente.'
  'Tip.Explorer' = 'Abrir el elemento seleccionado en el Explorador'
  'Tip.LogFolder' = 'Abrir la carpeta de registros (logs\)'
  'Tip.Export' = 'Guardar los resultados como .txt / .csv / .json (el formato sigue a la lista de extensiones)'
  'Tip.Abort' = 'Cancelar todas las tareas y salir (Ctrl+Q). Sin más confirmación; los detalles quedan en el registro.'
  'Tip.Refresh' = 'Actualizar la lista de unidades (tras conectar un USB)'
  'Tip.ToggleDrives' = 'Alternar la selección de todas las unidades'
  'Tip.UiLang' = 'Cambia solo el idioma de esta interfaz; no afecta a ZW3D ni a los resultados. La elección se recuerda.'
  'Tip.ScanLang' = 'Versión de idioma de ZW3D usada al buscar: afecta a las palabras predefinidas y términos localizados. «Todos los idiomas» = sin restricción. No cambia la interfaz.'
  'Tip.UserFolder' = 'Limpiar la carpeta de usuario de ZW3D con las mismas cuatro acciones que el User Folder Manager oficial'
  'Drive.Available' = '({0} / {1} GB libres)'
  'Status.ReadyFull' = 'Listo · seleccione una unidad e introduzca una palabra clave'
  'Status.Scanning' = 'Analizando…'
  'Status.ScanProgress' = 'Analizando… {0} revisados, {1} coincidencias, {2} carpetas sin acceso omitidas'
  'Status.Found' = '{0} encontrados ({1} analizados, {2} carpetas sin acceso omitidas)'
  'Status.Truncated' = ' · Límite de 5000 resultados alcanzado, reduzca el alcance'
  'Status.ExtHint' = ' · Consejo: {0} no está en el filtro ZW3D; desmarque «Solo archivos ZW3D»'
  'Status.Checked' = '{0} elemento(s) marcado(s)'
  'Status.Unchecked' = 'Selección borrada'
  'Status.DrivesRefreshed' = 'Lista de unidades actualizada'
  'Status.Exported' = '{0} fila(s) exportada(s) -> {1}'
  'Status.NoDelete' = 'Nada que eliminar'
  'Status.Cancelled' = 'Eliminación cancelada'
  'Status.SearchFail' = 'No se pudo iniciar la búsqueda, consulte el registro'
  'Status.Aborting' = 'Detención: cancelando todas las tareas…'
  'Status.Deleting' = 'Eliminando {0}/{1}: {2}'
  'Status.Summary' = '{0} correctos, {1} fallidos ({2})'
  'Msg.InfoTitle' = 'Información'
  'Msg.ErrTitle' = 'Error'
  'Msg.WarnTitle' = 'Advertencia'
  'Msg.NoDrive' = 'Seleccione al menos una unidad.'
  'Msg.NoExport' = 'No hay resultados para exportar. Realice primero una búsqueda.'
  'Msg.NoSelection' = 'Marque primero los elementos que desea eliminar.'
  'Msg.ProtectedTitle' = 'Rutas protegidas omitidas'
  'Msg.ProtectedBody' = 'Se omitieron las siguientes rutas protegidas del sistema:¶'
  'Msg.ExportOkTitle' = 'Exportación completada'
  'Msg.ExportOk' = 'Se exportaron {0} fila(s) a:¶{1}'
  'Msg.ExportFailBody' = 'Error al exportar:¶'
  'Msg.PartialTitle' = 'Algunos elementos no se pudieron eliminar'
  'Msg.ResultTitle' = 'Completado'
  'Msg.ResultBody' = 'Eliminación completada.¶¶{0}'
  'Msg.FailListBody' = 'Resultado: {0}¶¶Elementos fallidos:¶'
  'Msg.RecycleFail' = 'Componente de papelera no disponible; marque «Eliminar permanentemente» y vuelva a intentarlo'
  'Msg.ErrPermission' = 'Permiso denegado (UnauthorizedAccess)'
  'Msg.ErrSecurity' = 'Restricción de seguridad (SecurityException)'
  'Msg.ErrLocked' = 'Archivo / carpeta en uso por otro proceso: '
  'Msg.ErrLongPath' = 'Ruta demasiado larga o no válida: '
  'Msg.ErrIO' = 'Error de E/S: '
  'Msg.ErrUnexpected' = 'Error inesperado: {0}'
  'Dlg.Title' = 'Confirmar eliminación'
  'Dlg.Heading' = 'Se eliminarán {0} elemento(s)'
  'Dlg.Perm' = 'Modo: eliminación permanente. Los archivos no pasan por la papelera y no se pueden recuperar.'
  'Dlg.Recycle' = 'Modo: mover a la papelera. Se pueden restaurar.'
  'Dlg.Confirm' = 'Confirmo que deseo eliminar los elementos anteriores'
  'Dlg.Cancel' = 'Cancelar'
  'Dlg.Ok' = 'Eliminar'
  'Export.Title' = 'Exportar resultados de búsqueda'
  'Export.FilterTxt' = 'Archivo de texto (*.txt)|*.txt'
  'Export.FileName' = 'Lista_ZW3D_'
  'Mode.Permanent' = 'eliminación permanente'
  'Mode.Recycle' = 'movido a la papelera'
  'ZwMode.Off' = 'desactivado'
  'ZwMode.All' = 'todos los archivos ZW3D'
  'ZwMode.Junk' = 'solo temporales ZW3D'
  'ZwMode.Model' = 'solo modelos ZW3D'
  'Dlg.UfTitle' = 'Limpieza de la carpeta de usuario ZW3D'
  'Dlg.UfVersion' = 'Versión a limpiar:'
  'Dlg.UfLocation' = 'Ubicación:'
  'Dlg.UfHint' = 'Las versiones de ZW3D instaladas se detectan automáticamente; al cambiar de versión se actualizan las rutas.'
  'Dlg.UfOpen' = 'Abrir'
  'Dlg.UfClose' = 'Cerrar'
  'Dlg.UfNote' = 'Nota: la limpieza afecta a ZW3D en ejecución; ciérrelo primero. Se pedirá confirmación antes de eliminar.'
  'Uf.StatLine' = '{0} archivo(s) / {1} carpeta(s) / {2}'
  'Uf.NoTarget' = '(la carpeta no existe, nada que limpiar)'
  'Uf.ClearOutput' = 'Borrar archivos de salida'
  'Uf.ClearCustom' = 'Borrar configuración'
  'Uf.ClearTemp' = 'Borrar temporales'
  'Uf.ClearAll' = 'Vaciar carpeta de usuario'
  'Uf.DescOutput' = 'Elimina output\ (registros, archivos, conversiones)'
  'Uf.DescCustom' = 'Elimina custom\ (configuración de interfaz, plantillas, licencias – restablece ajustes personales)'
  'Uf.DescTemp' = 'Elimina temp\ (archivos de sesión e intermedios)'
  'Uf.DescAll' = 'Elimina toda la carpeta de la versión (todo lo anterior y más)'
  'Uf.Confirm' = 'Acción: {0}¶Versión: {5}¶Carpeta: {1}¶{2} archivo(s), {3} subcarpeta(s).¶Modo: {4}'
  'Uf.WarnCustom' = '¶¶Advertencia: se eliminarán las configuraciones personales; ZW3D las recreará en el próximo inicio.'
  'Uf.WarnAll' = '¶¶Advertencia: se vaciará toda la carpeta de usuario. Haga antes una copia de lo que necesite.'
  'Uf.ZwRunning' = 'ZW3D está en ejecución; puede haber archivos bloqueados. Cierre ZW3D primero. ¿Continuar de todos modos?'
  'Uf.NoFolder' = 'No se encontró la carpeta de usuario de ZW3D: {0}'
  'Uf.Missing' = 'La carpeta no existe, nada que limpiar: {0}'
  'Uf.Ok' = 'Limpieza completada: {0}'
  'Uf.Fail' = 'Error en la limpieza: {0}'
  'Uf.Cancelled' = 'Cancelado'
  'Uf.ConfirmTitle' = 'Confirmar limpieza'
 }
 'it_IT' = @{
  'Lang.All' = 'Tutte le lingue'
  'Hint' = 'Suggerimento: parola chiave vuota = elenca tutto (max 5000 risultati). Invio per cercare.'
  'Zw3dHint' = 'Criteri: elenco estensioni / nome con ZW3D·TopologicRep·Proxy / percorso con ZW3D'
  'ZwScope.All' = 'Tutti i file ZW3D (modelli + temp)'
  'ZwScope.Junk' = 'Solo file temporanei ZW3D (clear ALL.bat)'
  'ZwScope.Model' = 'Solo modelli ZW3D (.z3/.z3prt/.z3asm/.z3drw)'
  'CreditTip' = 'Sviluppatore'
  'Tip.Search' = 'Avvia la ricerca (o premi Invio nella casella della parola chiave)'
  'Tip.Stop' = 'Interrompi la scansione in corso'
  'Tip.CheckAll' = 'Seleziona tutti i risultati (pronti per l''eliminazione)'
  'Tip.Uncheck' = 'Rimuovi tutte le selezioni'
  'Tip.Delete' = 'Elimina gli elementi selezionati. Per impostazione predefinita vanno nel cestino (ripristinabili); attiva «Eliminazione definitiva» per cancellarli subito.'
  'Tip.Explorer' = 'Apri l''elemento selezionato in Esplora risorse'
  'Tip.LogFolder' = 'Apri la cartella dei log (logs\)'
  'Tip.Export' = 'Salva i risultati come .txt / .csv / .json (il formato segue la casella estensione)'
  'Tip.Abort' = 'Interrompe tutte le operazioni ed esce (Ctrl+Q). Nessuna ulteriore conferma; i dettagli sono nel log.'
  'Tip.Refresh' = 'Aggiorna l''elenco delle unità (dopo aver collegato una USB)'
  'Tip.ToggleDrives' = 'Inverti la selezione di tutte le unità'
  'Tip.UiLang' = 'Cambia solo la lingua di questa interfaccia; non modifica ZW3D né i risultati. La scelta viene ricordata.'
  'Tip.ScanLang' = 'Versione linguistica di ZW3D usata nella scansione: influisce sulle parole predefinite e sui termini localizzati. «Tutte le lingue» = nessun limite. Non cambia l''interfaccia.'
  'Tip.UserFolder' = 'Pulisci la cartella utente di ZW3D con le stesse quattro azioni dell''User Folder Manager ufficiale'
  'Drive.Available' = '({0} / {1} GB liberi)'
  'Status.ReadyFull' = 'Pronto · seleziona un''unità e inserisci una parola chiave'
  'Status.Scanning' = 'Scansione…'
  'Status.ScanProgress' = 'Scansione… {0} controllati, {1} trovati, {2} cartelle inaccessibili saltate'
  'Status.Found' = '{0} trovati ({1} analizzati, {2} cartelle inaccessibili saltate)'
  'Status.Truncated' = ' · Raggiunto il limite di 5000 risultati, restringi la ricerca'
  'Status.ExtHint' = ' · Suggerimento: {0} non è nel filtro ZW3D; deseleziona «Solo file ZW3D»'
  'Status.Checked' = '{0} elemento/i selezionato/i'
  'Status.Unchecked' = 'Selezione rimossa'
  'Status.DrivesRefreshed' = 'Elenco unità aggiornato'
  'Status.Exported' = '{0} riga/e esportata/e -> {1}'
  'Status.NoDelete' = 'Niente da eliminare'
  'Status.Cancelled' = 'Eliminazione annullata'
  'Status.SearchFail' = 'Impossibile avviare la ricerca, vedi il log'
  'Status.Aborting' = 'Arresto: interruzione di tutte le operazioni…'
  'Status.Deleting' = 'Eliminazione {0}/{1}: {2}'
  'Status.Summary' = '{0} riusciti, {1} non riusciti ({2})'
  'Msg.InfoTitle' = 'Informazione'
  'Msg.ErrTitle' = 'Errore'
  'Msg.WarnTitle' = 'Avviso'
  'Msg.NoDrive' = 'Seleziona almeno un''unità.'
  'Msg.NoExport' = 'Nessun risultato da esportare. Esegui prima una ricerca.'
  'Msg.NoSelection' = 'Seleziona prima gli elementi da eliminare.'
  'Msg.ProtectedTitle' = 'Percorsi protetti saltati'
  'Msg.ProtectedBody' = 'I seguenti percorsi di sistema protetti sono stati saltati:¶'
  'Msg.ExportOkTitle' = 'Esportazione completata'
  'Msg.ExportOk' = '{0} riga/e esportata/e in:¶{1}'
  'Msg.ExportFailBody' = 'Esportazione non riuscita:¶'
  'Msg.PartialTitle' = 'Alcuni elementi non sono stati eliminati'
  'Msg.ResultTitle' = 'Completato'
  'Msg.ResultBody' = 'Eliminazione completata.¶¶{0}'
  'Msg.FailListBody' = 'Risultato: {0}¶¶Elementi non riusciti:¶'
  'Msg.RecycleFail' = 'Componente cestino non disponibile; attiva «Eliminazione definitiva» e riprova'
  'Msg.ErrPermission' = 'Accesso negato (UnauthorizedAccess)'
  'Msg.ErrSecurity' = 'Restrizione di sicurezza (SecurityException)'
  'Msg.ErrLocked' = 'File / cartella in uso da un altro processo: '
  'Msg.ErrLongPath' = 'Percorso troppo lungo o non valido: '
  'Msg.ErrIO' = 'Errore di I/O: '
  'Msg.ErrUnexpected' = 'Errore imprevisto: {0}'
  'Dlg.Title' = 'Conferma eliminazione'
  'Dlg.Heading' = 'Stai per eliminare {0} elemento/i'
  'Dlg.Perm' = 'Modalità: eliminazione definitiva. I file non passano dal cestino e non possono essere ripristinati.'
  'Dlg.Recycle' = 'Modalità: sposta nel cestino. Gli elementi possono essere ripristinati.'
  'Dlg.Confirm' = 'Confermo di voler eliminare gli elementi sopra'
  'Dlg.Cancel' = 'Annulla'
  'Dlg.Ok' = 'Elimina'
  'Export.Title' = 'Esporta risultati della ricerca'
  'Export.FilterTxt' = 'File di testo (*.txt)|*.txt'
  'Export.FileName' = 'Elenco_ZW3D_'
  'Mode.Permanent' = 'eliminazione definitiva'
  'Mode.Recycle' = 'spostato nel cestino'
  'ZwMode.Off' = 'disattivato'
  'ZwMode.All' = 'tutti i file ZW3D'
  'ZwMode.Junk' = 'solo temporanei ZW3D'
  'ZwMode.Model' = 'solo modelli ZW3D'
  'Dlg.UfTitle' = 'Pulizia cartella utente ZW3D'
  'Dlg.UfVersion' = 'Versione da pulire:'
  'Dlg.UfLocation' = 'Posizione:'
  'Dlg.UfHint' = 'Le versioni di ZW3D installate vengono rilevate automaticamente; cambiando versione i percorsi si aggiornano.'
  'Dlg.UfOpen' = 'Apri'
  'Dlg.UfClose' = 'Chiudi'
  'Dlg.UfNote' = 'Nota: la pulizia influisce su ZW3D in esecuzione; chiudilo prima. Verrà chiesta conferma prima di eliminare.'
  'Uf.StatLine' = '{0} file / {1} cartelle / {2}'
  'Uf.NoTarget' = '(cartella inesistente, niente da pulire)'
  'Uf.ClearOutput' = 'Cancella file di output'
  'Uf.ClearCustom' = 'Cancella configurazioni'
  'Uf.ClearTemp' = 'Cancella file temporanei'
  'Uf.ClearAll' = 'Svuota cartella utente'
  'Uf.DescOutput' = 'Elimina output\ (log, archivi, file convertiti)'
  'Uf.DescCustom' = 'Elimina custom\ (configurazione UI, modelli, licenze – ripristina le impostazioni personali)'
  'Uf.DescTemp' = 'Elimina temp\ (file di sessione e intermedi)'
  'Uf.DescAll' = 'Elimina l''intera cartella della versione (tutto quanto sopra e altro)'
  'Uf.Confirm' = 'Azione: {0}¶Versione: {5}¶Cartella: {1}¶{2} file, {3} sottocartelle.¶Modalità: {4}'
  'Uf.WarnCustom' = '¶¶Attenzione: le configurazioni personali verranno rimosse; ZW3D le ricrea al prossimo avvio.'
  'Uf.WarnAll' = '¶¶Attenzione: l''intera cartella utente verrà svuotata. Esegui prima un backup di ciò che ti serve.'
  'Uf.ZwRunning' = 'ZW3D è in esecuzione; alcuni file potrebbero essere bloccati. Chiudi prima ZW3D. Continuare comunque?'
  'Uf.NoFolder' = 'Cartella utente ZW3D non trovata: {0}'
  'Uf.Missing' = 'La cartella non esiste, niente da pulire: {0}'
  'Uf.Ok' = 'Pulizia completata: {0}'
  'Uf.Fail' = 'Pulizia non riuscita: {0}'
  'Uf.Cancelled' = 'Annullato'
  'Uf.ConfirmTitle' = 'Conferma pulizia'
 }
 'pt_PT' = @{
  'Lang.All' = 'Todos os idiomas'
  'Hint' = 'Dica: palavra-chave vazia = listar tudo (máx. 5000 resultados). Enter para pesquisar.'
  'Zw3dHint' = 'Critério: lista de extensões / nome com ZW3D·TopologicRep·Proxy / caminho com ZW3D'
  'ZwScope.All' = 'Todos os ficheiros ZW3D (modelos + temp.)'
  'ZwScope.Junk' = 'Apenas temporários ZW3D (clear ALL.bat)'
  'ZwScope.Model' = 'Apenas modelos ZW3D (.z3/.z3prt/.z3asm/.z3drw)'
  'CreditTip' = 'Programador'
  'Tip.Search' = 'Iniciar pesquisa (ou prima Enter na caixa de palavra-chave)'
  'Tip.Stop' = 'Parar a análise em curso'
  'Tip.CheckAll' = 'Marcar todos os resultados (prontos a eliminar)'
  'Tip.Uncheck' = 'Limpar todas as marcas'
  'Tip.Delete' = 'Eliminar os itens marcados. Por predefinição vão para a Reciclagem (recuperáveis); marque «Eliminar definitivamente» para os apagar diretamente.'
  'Tip.Explorer' = 'Abrir o item selecionado no Explorador'
  'Tip.LogFolder' = 'Abrir a pasta de registos (logs\)'
  'Tip.Export' = 'Guardar os resultados como .txt / .csv / .json (formato segue a caixa de extensão)'
  'Tip.Abort' = 'Interromper todas as tarefas e sair (Ctrl+Q). Sem mais confirmação; os detalhes ficam no registo.'
  'Tip.Refresh' = 'Atualizar a lista de unidades (depois de ligar uma USB)'
  'Tip.ToggleDrives' = 'Alternar a seleção de todas as unidades'
  'Tip.UiLang' = 'Muda apenas o idioma desta interface; não altera o ZW3D nem os resultados. A escolha é memorizada.'
  'Tip.ScanLang' = 'Versão de idioma do ZW3D usada na análise: afeta as palavras predefinidas e termos localizados. «Todos os idiomas» = sem restrição. Não altera a interface.'
  'Tip.UserFolder' = 'Limpar a pasta de utilizador do ZW3D com as mesmas quatro ações do User Folder Manager oficial'
  'Drive.Available' = '({0} / {1} GB livres)'
  'Status.ReadyFull' = 'Pronto · selecione uma unidade e introduza uma palavra-chave'
  'Status.Scanning' = 'A analisar…'
  'Status.ScanProgress' = 'A analisar… {0} verificados, {1} encontrados, {2} pastas sem acesso ignoradas'
  'Status.Found' = '{0} encontrados ({1} analisados, {2} pastas sem acesso ignoradas)'
  'Status.Truncated' = ' · Limite de 5000 resultados atingido, restrinja a pesquisa'
  'Status.ExtHint' = ' · Dica: {0} não está no filtro ZW3D; desmarque «Apenas ficheiros ZW3D»'
  'Status.Checked' = '{0} item(ns) marcado(s)'
  'Status.Unchecked' = 'Seleção limpa'
  'Status.DrivesRefreshed' = 'Lista de unidades atualizada'
  'Status.Exported' = '{0} linha(s) exportada(s) -> {1}'
  'Status.NoDelete' = 'Nada para eliminar'
  'Status.Cancelled' = 'Eliminação cancelada'
  'Status.SearchFail' = 'Não foi possível iniciar a pesquisa, consulte o registo'
  'Status.Aborting' = 'Paragem: a interromper todas as tarefas…'
  'Status.Deleting' = 'A eliminar {0}/{1}: {2}'
  'Status.Summary' = '{0} com êxito, {1} falhados ({2})'
  'Msg.InfoTitle' = 'Informação'
  'Msg.ErrTitle' = 'Erro'
  'Msg.WarnTitle' = 'Aviso'
  'Msg.NoDrive' = 'Selecione pelo menos uma unidade.'
  'Msg.NoExport' = 'Não há resultados para exportar. Faça primeiro uma pesquisa.'
  'Msg.NoSelection' = 'Marque primeiro os itens a eliminar.'
  'Msg.ProtectedTitle' = 'Caminhos protegidos ignorados'
  'Msg.ProtectedBody' = 'Os seguintes caminhos de sistema protegidos foram ignorados:¶'
  'Msg.ExportOkTitle' = 'Exportação concluída'
  'Msg.ExportOk' = '{0} linha(s) exportada(s) para:¶{1}'
  'Msg.ExportFailBody' = 'Falha na exportação:¶'
  'Msg.PartialTitle' = 'Não foi possível eliminar alguns itens'
  'Msg.ResultTitle' = 'Concluído'
  'Msg.ResultBody' = 'Eliminação concluída.¶¶{0}'
  'Msg.FailListBody' = 'Resultado: {0}¶¶Itens falhados:¶'
  'Msg.RecycleFail' = 'Componente da Reciclagem indisponível; marque «Eliminar definitivamente» e tente de novo'
  'Msg.ErrPermission' = 'Acesso negado (UnauthorizedAccess)'
  'Msg.ErrSecurity' = 'Restrição de segurança (SecurityException)'
  'Msg.ErrLocked' = 'Ficheiro / pasta em uso por outro processo: '
  'Msg.ErrLongPath' = 'Caminho demasiado longo ou inválido: '
  'Msg.ErrIO' = 'Erro de E/S: '
  'Msg.ErrUnexpected' = 'Erro inesperado: {0}'
  'Dlg.Title' = 'Confirmar eliminação'
  'Dlg.Heading' = 'Vão ser eliminados {0} item(ns)'
  'Dlg.Perm' = 'Modo: eliminação definitiva. Os ficheiros não passam pela Reciclagem e não podem ser recuperados.'
  'Dlg.Recycle' = 'Modo: mover para a Reciclagem. Os itens podem ser restaurados.'
  'Dlg.Confirm' = 'Confirmo que pretendo eliminar os itens acima'
  'Dlg.Cancel' = 'Cancelar'
  'Dlg.Ok' = 'Eliminar'
  'Export.Title' = 'Exportar resultados da pesquisa'
  'Export.FilterTxt' = 'Ficheiro de texto (*.txt)|*.txt'
  'Export.FileName' = 'Lista_ZW3D_'
  'Mode.Permanent' = 'eliminação definitiva'
  'Mode.Recycle' = 'movido para a Reciclagem'
  'ZwMode.Off' = 'desligado'
  'ZwMode.All' = 'todos os ficheiros ZW3D'
  'ZwMode.Junk' = 'apenas temporários ZW3D'
  'ZwMode.Model' = 'apenas modelos ZW3D'
  'Dlg.UfTitle' = 'Limpeza da pasta de utilizador ZW3D'
  'Dlg.UfVersion' = 'Versão a limpar:'
  'Dlg.UfLocation' = 'Localização:'
  'Dlg.UfHint' = 'As versões do ZW3D instaladas são detetadas automaticamente; ao mudar de versão os caminhos são atualizados.'
  'Dlg.UfOpen' = 'Abrir'
  'Dlg.UfClose' = 'Fechar'
  'Dlg.UfNote' = 'Nota: a limpeza afeta o ZW3D em execução; feche-o primeiro. É pedida confirmação antes de eliminar.'
  'Uf.StatLine' = '{0} ficheiro(s) / {1} pasta(s) / {2}'
  'Uf.NoTarget' = '(a pasta não existe, nada a limpar)'
  'Uf.ClearOutput' = 'Limpar ficheiros de saída'
  'Uf.ClearCustom' = 'Limpar configurações'
  'Uf.ClearTemp' = 'Limpar temporários'
  'Uf.ClearAll' = 'Esvaziar pasta de utilizador'
  'Uf.DescOutput' = 'Elimina output\ (registos, arquivos, ficheiros convertidos)'
  'Uf.DescCustom' = 'Elimina custom\ (configuração da interface, modelos, licenças – repõe as definições pessoais)'
  'Uf.DescTemp' = 'Elimina temp\ (ficheiros de sessão e intermédios)'
  'Uf.DescAll' = 'Elimina toda a pasta da versão (tudo o que está acima e mais)'
  'Uf.Confirm' = 'Ação: {0}¶Versão: {5}¶Pasta: {1}¶{2} ficheiro(s), {3} subpasta(s).¶Modo: {4}'
  'Uf.WarnCustom' = '¶¶Aviso: as configurações pessoais serão removidas; o ZW3D recria-as no próximo arranque.'
  'Uf.WarnAll' = '¶¶Aviso: toda a pasta de utilizador será esvaziada. Faça primeiro uma cópia do que precisar.'
  'Uf.ZwRunning' = 'O ZW3D está em execução; os ficheiros podem estar bloqueados. Feche primeiro o ZW3D. Continuar mesmo assim?'
  'Uf.NoFolder' = 'Pasta de utilizador do ZW3D não encontrada: {0}'
  'Uf.Missing' = 'A pasta não existe, nada a limpar: {0}'
  'Uf.Ok' = 'Limpeza concluída: {0}'
  'Uf.Fail' = 'Falha na limpeza: {0}'
  'Uf.Cancelled' = 'Cancelado'
  'Uf.ConfirmTitle' = 'Confirmar limpeza'
 }
 'ja_JP' = @{
  'Lang.All' = 'すべての言語'
  'Hint' = 'ヒント：キーワード空欄 = すべて表示（最大 5000 件）。Enter で検索。'
  'Zw3dHint' = '判定：拡張子リスト／名前に ZW3D·TopologicRep·Proxy を含む／パスに ZW3D を含む'
  'ZwScope.All' = 'すべての ZW3D ファイル（モデル＋一時）'
  'ZwScope.Junk' = 'ZW3D 一時ファイルのみ（clear ALL.bat 相当）'
  'ZwScope.Model' = 'ZW3D モデルのみ（.z3/.z3prt/.z3asm/.z3drw）'
  'CreditTip' = '開発者'
  'Tip.Search' = '検索を開始（キーワード欄で Enter でも可）'
  'Tip.Stop' = '実行中のスキャンを停止'
  'Tip.CheckAll' = 'すべての結果にチェック（削除準備）'
  'Tip.Uncheck' = 'すべてのチェックを外す'
  'Tip.Delete' = 'チェックした項目を削除します。既定はごみ箱へ移動（復元可）。上の「完全に削除」をオンにすると直接削除します。'
  'Tip.Explorer' = '選択した項目をエクスプローラーで開く'
  'Tip.LogFolder' = 'ログフォルダー（logs\）を開く'
  'Tip.Export' = '検索結果を .txt / .csv / .json で保存（形式は拡張子欄に従う）'
  'Tip.Abort' = 'すべての処理を中止して終了（Ctrl+Q）。確認なし。詳細はログに記録されます。'
  'Tip.Refresh' = 'ドライブ一覧を再取得（USB 接続後に使用）'
  'Tip.ToggleDrives' = 'すべてのドライブのチェックを切り替え'
  'Tip.UiLang' = 'このツールの表示言語だけを切り替えます。ZW3D 本体や検索結果には影響しません。選択は保存されます。'
  'Tip.ScanLang' = 'スキャン時に識別する ZW3D の言語版：よく使うキーワードとローカライズ語に影響します。「すべての言語」= 制限なし。表示言語は変わりません。'
  'Tip.UserFolder' = '公式 User Folder Manager と同じ 4 つの操作で ZW3D ユーザーフォルダーを整理'
  'Drive.Available' = '（{0} / {1} GB 空き）'
  'Status.ReadyFull' = '準備完了 · ドライブを選んでキーワードを入力'
  'Status.Scanning' = 'スキャン中…'
  'Status.ScanProgress' = 'スキャン中… {0} 件確認、{1} 件一致、アクセス不可フォルダー {2} 件をスキップ'
  'Status.Found' = '{0} 件見つかりました（{1} 件スキャン、アクセス不可フォルダー {2} 件をスキップ）'
  'Status.Truncated' = ' · 上限 5000 件に達しました。範囲を絞ってください'
  'Status.ExtHint' = ' · ヒント：{0} は ZW3D フィルターにありません。「ZW3D 関連のみ」を外してください'
  'Status.Checked' = '{0} 件をチェック'
  'Status.Unchecked' = 'チェックを解除しました'
  'Status.DrivesRefreshed' = 'ドライブ一覧を更新しました'
  'Status.Exported' = '{0} 件を出力 → {1}'
  'Status.NoDelete' = '削除する項目がありません'
  'Status.Cancelled' = '削除を取り消しました'
  'Status.SearchFail' = '検索を開始できません。ログを確認してください'
  'Status.Aborting' = '即時終了：すべての処理を中止しています…'
  'Status.Deleting' = '削除中 {0}/{1}：{2}'
  'Status.Summary' = '成功 {0} 件、失敗 {1} 件（{2}）'
  'Msg.InfoTitle' = 'お知らせ'
  'Msg.ErrTitle' = 'エラー'
  'Msg.WarnTitle' = '警告'
  'Msg.NoDrive' = 'ドライブを 1 つ以上選択してください。'
  'Msg.NoExport' = '出力できる結果がありません。先に検索してください。'
  'Msg.NoSelection' = '削除する項目に先にチェックを入れてください。'
  'Msg.ProtectedTitle' = '保護パスをスキップしました'
  'Msg.ProtectedBody' = '次のシステム保護パスは自動的にスキップされました：¶'
  'Msg.ExportOkTitle' = '出力完了'
  'Msg.ExportOk' = '{0} 件を出力しました：¶{1}'
  'Msg.ExportFailBody' = '出力に失敗しました：¶'
  'Msg.PartialTitle' = '一部の項目を削除できませんでした'
  'Msg.ResultTitle' = '完了'
  'Msg.ResultBody' = '削除が完了しました。¶¶{0}'
  'Msg.FailListBody' = '削除結果：{0}¶¶失敗した項目：¶'
  'Msg.RecycleFail' = 'ごみ箱コンポーネントを利用できません。「完全に削除」をオンにして再試行してください'
  'Msg.ErrPermission' = 'アクセス権がありません（UnauthorizedAccess）'
  'Msg.ErrSecurity' = 'セキュリティ制限（SecurityException）'
  'Msg.ErrLocked' = 'ファイル／フォルダーが別のプロセスで使用中：'
  'Msg.ErrLongPath' = 'パスが長すぎるか無効です：'
  'Msg.ErrIO' = 'I/O エラー：'
  'Msg.ErrUnexpected' = '予期しないエラー：{0}'
  'Dlg.Title' = '削除の確認'
  'Dlg.Heading' = '{0} 個の項目を削除します'
  'Dlg.Perm' = 'モード：完全に削除。ごみ箱を経由せず、削除後は復元できません。'
  'Dlg.Recycle' = 'モード：ごみ箱へ移動。ごみ箱から復元できます。'
  'Dlg.Confirm' = '上記の項目を削除することを確認しました'
  'Dlg.Cancel' = 'キャンセル'
  'Dlg.Ok' = '削除する'
  'Export.Title' = '検索結果の一覧を出力'
  'Export.FilterTxt' = 'テキストファイル (*.txt)|*.txt'
  'Export.FileName' = 'ZW3D一覧_'
  'Mode.Permanent' = '完全に削除'
  'Mode.Recycle' = 'ごみ箱へ移動'
  'ZwMode.Off' = 'オフ'
  'ZwMode.All' = 'すべての ZW3D ファイル'
  'ZwMode.Junk' = 'ZW3D 一時ファイルのみ'
  'ZwMode.Model' = 'ZW3D モデルのみ'
  'Dlg.UfTitle' = 'ZW3D ユーザーフォルダーの整理'
  'Dlg.UfVersion' = '削除するバージョン：'
  'Dlg.UfLocation' = '場所：'
  'Dlg.UfHint' = 'インストール済みの ZW3D バージョンを自動検出します。切り替えるとパスも自動更新されます。'
  'Dlg.UfOpen' = '開く'
  'Dlg.UfClose' = '閉じる'
  'Dlg.UfNote' = '注意：実行中の ZW3D に影響します。先に終了してください。削除前に再確認します。'
  'Uf.StatLine' = 'ファイル {0} 個 · フォルダー {1} 個 · {2}'
  'Uf.NoTarget' = '（フォルダーが存在しないため整理不要）'
  'Uf.ClearOutput' = '自動出力ファイルを削除'
  'Uf.ClearCustom' = 'カスタム設定を削除'
  'Uf.ClearTemp' = '一時ファイルを削除'
  'Uf.ClearAll' = 'ユーザーフォルダーを削除'
  'Uf.DescOutput' = 'output\ を削除（ログ、アーカイブ、変換出力など）'
  'Uf.DescCustom' = 'custom\ を削除（UI 設定、テンプレート、ライセンスなど。個人設定がリセットされます）'
  'Uf.DescTemp' = 'temp\ を削除（セッションと中間一時ファイル）'
  'Uf.DescAll' = 'バージョンフォルダー全体を削除（上記すべてとその他のファイル）'
  'Uf.Confirm' = '実行内容：{0}¶バージョン：{5}¶フォルダー：{1}¶ファイル {2} 個、サブフォルダー {3} 個。¶モード：{4}'
  'Uf.WarnCustom' = '¶¶警告：個人設定が削除されます。ZW3D は次回起動時に再作成します。'
  'Uf.WarnAll' = '¶¶警告：ユーザーフォルダー全体が空になります。必要な設定は先にバックアップしてください。'
  'Uf.ZwRunning' = 'ZW3D が実行中です。ファイルがロックされている可能性があります。先に ZW3D を終了することをお勧めします。続行しますか？'
  'Uf.NoFolder' = 'ZW3D ユーザーフォルダーが見つかりません：{0}'
  'Uf.Missing' = 'フォルダーが存在しないため整理不要：{0}'
  'Uf.Ok' = '整理完了：{0}'
  'Uf.Fail' = '整理に失敗しました：{0}'
  'Uf.Cancelled' = 'キャンセルしました'
  'Uf.ConfirmTitle' = '整理の確認'
 }
 'ko_KR' = @{
  'Lang.All' = '모든 언어'
  'Hint' = '팁: 키워드를 비우면 전체 표시(최대 5000개). Enter로 바로 검색.'
  'Zw3dHint' = '판정: 확장자 목록 / 이름에 ZW3D·TopologicRep·Proxy 포함 / 경로에 ZW3D 포함'
  'ZwScope.All' = '모든 ZW3D 파일(모델 + 임시)'
  'ZwScope.Junk' = 'ZW3D 임시 파일만(clear ALL.bat 대응)'
  'ZwScope.Model' = 'ZW3D 모델만(.z3/.z3prt/.z3asm/.z3drw)'
  'CreditTip' = '개발자'
  'Tip.Search' = '검색 시작(키워드 칸에서 Enter를 눌러도 됨)'
  'Tip.Stop' = '진행 중인 검사 중지'
  'Tip.CheckAll' = '모든 결과에 체크(삭제 준비)'
  'Tip.Uncheck' = '모든 체크 해제'
  'Tip.Delete' = '체크한 항목을 삭제합니다. 기본은 휴지통으로 이동(복원 가능)이며, 위의 "영구 삭제"를 체크하면 바로 삭제합니다.'
  'Tip.Explorer' = '선택한 항목을 탐색기에서 열기'
  'Tip.LogFolder' = '로그 폴더(logs\) 열기'
  'Tip.Export' = '검색 결과를 .txt / .csv / .json으로 저장(형식은 확장자 칸을 따름)'
  'Tip.Abort' = '모든 작업을 중단하고 종료(Ctrl+Q). 다시 묻지 않으며 자세한 내용은 로그에 기록됩니다.'
  'Tip.Refresh' = '드라이브 목록 다시 읽기(USB 연결 후 사용)'
  'Tip.ToggleDrives' = '모든 드라이브의 체크 상태 전환'
  'Tip.UiLang' = '이 도구의 표시 언어만 바꿉니다. ZW3D 자체나 검색 결과에는 영향이 없습니다. 선택은 저장됩니다.'
  'Tip.ScanLang' = '검사 시 인식할 ZW3D 언어 버전: 자주 쓰는 키워드와 현지화 검색어에 영향을 줍니다. "모든 언어" = 제한 없음. 표시 언어는 바뀌지 않습니다.'
  'Tip.UserFolder' = '공식 User Folder Manager와 같은 네 가지 작업으로 ZW3D 사용자 폴더 정리'
  'Drive.Available' = '({0} / {1} GB 사용 가능)'
  'Status.ReadyFull' = '준비 · 드라이브를 선택하고 키워드를 입력하세요'
  'Status.Scanning' = '검사 중…'
  'Status.ScanProgress' = '검사 중… {0}개 확인, {1}개 일치, 접근 불가 폴더 {2}개 건너뜀'
  'Status.Found' = '{0}개 찾음({1}개 검사, 접근 불가 폴더 {2}개 건너뜀)'
  'Status.Truncated' = ' · 결과 상한 5000개에 도달했습니다. 범위를 좁히세요'
  'Status.ExtHint' = ' · 팁: {0}은(는) ZW3D 필터 목록에 없습니다. "ZW3D 관련만"을 해제하세요'
  'Status.Checked' = '{0}개 체크됨'
  'Status.Unchecked' = '체크를 해제했습니다'
  'Status.DrivesRefreshed' = '드라이브 목록을 새로 고쳤습니다'
  'Status.Exported' = '{0}개 내보냄 → {1}'
  'Status.NoDelete' = '삭제할 항목이 없습니다'
  'Status.Cancelled' = '삭제를 취소했습니다'
  'Status.SearchFail' = '검색을 시작하지 못했습니다. 로그를 확인하세요'
  'Status.Aborting' = '즉시 중지: 모든 작업을 중단하는 중…'
  'Status.Deleting' = '삭제 중 {0}/{1}: {2}'
  'Status.Summary' = '성공 {0}개, 실패 {1}개({2})'
  'Msg.InfoTitle' = '알림'
  'Msg.ErrTitle' = '오류'
  'Msg.WarnTitle' = '경고'
  'Msg.NoDrive' = '드라이브를 하나 이상 선택하세요.'
  'Msg.NoExport' = '내보낼 결과가 없습니다. 먼저 검색하세요.'
  'Msg.NoSelection' = '삭제할 항목을 먼저 체크하세요.'
  'Msg.ProtectedTitle' = '보호 경로를 건너뜀'
  'Msg.ProtectedBody' = '다음 시스템 보호 경로는 자동으로 건너뛰었습니다:¶'
  'Msg.ExportOkTitle' = '내보내기 완료'
  'Msg.ExportOk' = '{0}개를 다음 위치로 내보냈습니다:¶{1}'
  'Msg.ExportFailBody' = '내보내기 실패:¶'
  'Msg.PartialTitle' = '일부 항목을 삭제하지 못했습니다'
  'Msg.ResultTitle' = '완료'
  'Msg.ResultBody' = '삭제가 완료되었습니다.¶¶{0}'
  'Msg.FailListBody' = '삭제 결과: {0}¶¶실패한 항목:¶'
  'Msg.RecycleFail' = '휴지통 구성 요소를 사용할 수 없습니다. "영구 삭제"를 체크한 후 다시 시도하세요'
  'Msg.ErrPermission' = '권한 부족(UnauthorizedAccess)'
  'Msg.ErrSecurity' = '보안 제한(SecurityException)'
  'Msg.ErrLocked' = '다른 프로그램이 파일/폴더를 사용 중: '
  'Msg.ErrLongPath' = '경로가 너무 길거나 잘못됨: '
  'Msg.ErrIO' = 'I/O 오류: '
  'Msg.ErrUnexpected' = '예기치 않은 오류: {0}'
  'Dlg.Title' = '삭제 확인'
  'Dlg.Heading' = '{0}개 항목을 삭제합니다'
  'Dlg.Perm' = '모드: 영구 삭제. 파일이 휴지통을 거치지 않으며 삭제 후 복원할 수 없습니다.'
  'Dlg.Recycle' = '모드: 휴지통으로 이동. 휴지통에서 복원할 수 있습니다.'
  'Dlg.Confirm' = '위 항목을 삭제하는 것에 동의합니다'
  'Dlg.Cancel' = '취소'
  'Dlg.Ok' = '삭제'
  'Export.Title' = '검색 결과 목록 내보내기'
  'Export.FilterTxt' = '텍스트 파일 (*.txt)|*.txt'
  'Export.FileName' = 'ZW3D목록_'
  'Mode.Permanent' = '영구 삭제'
  'Mode.Recycle' = '휴지통으로 이동'
  'ZwMode.Off' = '끄기'
  'ZwMode.All' = '모든 ZW3D 파일'
  'ZwMode.Junk' = 'ZW3D 임시 파일만'
  'ZwMode.Model' = 'ZW3D 모델만'
  'Dlg.UfTitle' = 'ZW3D 사용자 폴더 정리'
  'Dlg.UfVersion' = '정리할 버전:'
  'Dlg.UfLocation' = '위치:'
  'Dlg.UfHint' = '설치된 ZW3D 버전을 자동으로 감지합니다. 버전을 바꾸면 경로도 자동으로 갱신됩니다.'
  'Dlg.UfOpen' = '열기'
  'Dlg.UfClose' = '닫기'
  'Dlg.UfNote' = '주의: 정리는 실행 중인 ZW3D에 영향을 줍니다. 먼저 종료하세요. 삭제 전에 다시 확인합니다.'
  'Uf.StatLine' = '파일 {0}개 · 폴더 {1}개 · {2}'
  'Uf.NoTarget' = '(폴더가 없어 정리할 필요 없음)'
  'Uf.ClearOutput' = '자동 출력 파일 삭제'
  'Uf.ClearCustom' = '사용자 지정 구성 삭제'
  'Uf.ClearTemp' = '임시 파일 삭제'
  'Uf.ClearAll' = '사용자 폴더 삭제'
  'Uf.DescOutput' = 'output\ 삭제(로그, 보관, 변환 출력 등)'
  'Uf.DescCustom' = 'custom\ 삭제(UI 구성, 템플릿, 라이선스 등 — 개인 설정이 초기화됨)'
  'Uf.DescTemp' = 'temp\ 삭제(세션 및 중간 임시 파일)'
  'Uf.DescAll' = '버전 폴더 전체 삭제(위 항목 전부와 기타 파일 포함)'
  'Uf.Confirm' = '실행: {0}¶버전: {5}¶폴더: {1}¶파일 {2}개, 하위 폴더 {3}개.¶모드: {4}'
  'Uf.WarnCustom' = '¶¶경고: 개인 구성이 삭제됩니다. ZW3D는 다음 실행 시 다시 만듭니다.'
  'Uf.WarnAll' = '¶¶경고: 사용자 폴더 전체가 비워집니다. 필요한 설정은 먼저 백업하세요.'
  'Uf.ZwRunning' = 'ZW3D가 실행 중이며 파일이 잠겨 있을 수 있습니다. 먼저 ZW3D를 종료하는 것이 좋습니다. 계속하시겠습니까?'
  'Uf.NoFolder' = 'ZW3D 사용자 폴더를 찾을 수 없습니다: {0}'
  'Uf.Missing' = '폴더가 없어 정리할 필요 없음: {0}'
  'Uf.Ok' = '정리 완료: {0}'
  'Uf.Fail' = '정리 실패: {0}'
  'Uf.Cancelled' = '취소됨'
  'Uf.ConfirmTitle' = '정리 확인'
 }
 'ru_RU' = @{
  'Lang.All' = 'Все языки'
  'Hint' = 'Совет: пустое ключевое слово = показать всё (до 5000). Enter — поиск.'
  'Zw3dHint' = 'Критерии: список расширений / имя содержит ZW3D·TopologicRep·Proxy / путь содержит ZW3D'
  'ZwScope.All' = 'Все файлы ZW3D (модели + временные)'
  'ZwScope.Junk' = 'Только временные файлы ZW3D (clear ALL.bat)'
  'ZwScope.Model' = 'Только модели ZW3D (.z3/.z3prt/.z3asm/.z3drw)'
  'CreditTip' = 'Разработчик'
  'Tip.Search' = 'Начать поиск (или нажмите Enter в поле ключевого слова)'
  'Tip.Stop' = 'Остановить текущее сканирование'
  'Tip.CheckAll' = 'Отметить все результаты (для удаления)'
  'Tip.Uncheck' = 'Снять все отметки'
  'Tip.Delete' = 'Удалить отмеченные элементы. По умолчанию — в корзину (можно восстановить); включите «Удалить безвозвратно», чтобы стереть сразу.'
  'Tip.Explorer' = 'Открыть выбранный элемент в Проводнике'
  'Tip.LogFolder' = 'Открыть папку журналов (logs\)'
  'Tip.Export' = 'Сохранить результаты в .txt / .csv / .json (формат — по полю расширения)'
  'Tip.Abort' = 'Прервать все операции и выйти (Ctrl+Q). Без подтверждения; подробности — в журнале.'
  'Tip.Refresh' = 'Обновить список дисков (после подключения USB)'
  'Tip.ToggleDrives' = 'Переключить отметку всех дисков'
  'Tip.UiLang' = 'Меняет только язык интерфейса этой программы, не затрагивая ZW3D и результаты поиска. Выбор запоминается.'
  'Tip.ScanLang' = 'Языковая версия ZW3D для сканирования: влияет на шаблоны ключевых слов и локализованные термины. «Все языки» — без ограничений. Интерфейс не меняется.'
  'Tip.UserFolder' = 'Очистить папку пользователя ZW3D теми же четырьмя действиями, что и официальный User Folder Manager'
  'Drive.Available' = '({0} / {1} ГБ свободно)'
  'Status.ReadyFull' = 'Готово · выберите диск и введите ключевое слово'
  'Status.Scanning' = 'Сканирование…'
  'Status.ScanProgress' = 'Сканирование… проверено {0}, найдено {1}, пропущено недоступных папок: {2}'
  'Status.Found' = 'Найдено: {0} (проверено {1}, пропущено недоступных папок: {2})'
  'Status.Truncated' = ' · Достигнут лимит 5000 результатов, сузьте область поиска'
  'Status.ExtHint' = ' · Совет: {0} нет в списке фильтра ZW3D; снимите «Только файлы ZW3D»'
  'Status.Checked' = 'Отмечено: {0}'
  'Status.Unchecked' = 'Отметки сняты'
  'Status.DrivesRefreshed' = 'Список дисков обновлён'
  'Status.Exported' = 'Экспортировано строк: {0} -> {1}'
  'Status.NoDelete' = 'Нечего удалять'
  'Status.Cancelled' = 'Удаление отменено'
  'Status.SearchFail' = 'Не удалось запустить поиск, см. журнал'
  'Status.Aborting' = 'Останов: прерывание всех операций…'
  'Status.Deleting' = 'Удаление {0}/{1}: {2}'
  'Status.Summary' = 'Успешно: {0}, ошибок: {1} ({2})'
  'Msg.InfoTitle' = 'Информация'
  'Msg.ErrTitle' = 'Ошибка'
  'Msg.WarnTitle' = 'Предупреждение'
  'Msg.NoDrive' = 'Выберите хотя бы один диск.'
  'Msg.NoExport' = 'Нет результатов для экспорта. Сначала выполните поиск.'
  'Msg.NoSelection' = 'Сначала отметьте элементы для удаления.'
  'Msg.ProtectedTitle' = 'Защищённые пути пропущены'
  'Msg.ProtectedBody' = 'Следующие защищённые системные пути были пропущены:¶'
  'Msg.ExportOkTitle' = 'Экспорт завершён'
  'Msg.ExportOk' = 'Экспортировано строк: {0} в:¶{1}'
  'Msg.ExportFailBody' = 'Ошибка экспорта:¶'
  'Msg.PartialTitle' = 'Некоторые элементы не удалось удалить'
  'Msg.ResultTitle' = 'Готово'
  'Msg.ResultBody' = 'Удаление завершено.¶¶{0}'
  'Msg.FailListBody' = 'Результат: {0}¶¶Неудачные элементы:¶'
  'Msg.RecycleFail' = 'Компонент корзины недоступен; включите «Удалить безвозвратно» и повторите'
  'Msg.ErrPermission' = 'Доступ запрещён (UnauthorizedAccess)'
  'Msg.ErrSecurity' = 'Ограничение безопасности (SecurityException)'
  'Msg.ErrLocked' = 'Файл / папка используется другим процессом: '
  'Msg.ErrLongPath' = 'Путь слишком длинный или недопустимый: '
  'Msg.ErrIO' = 'Ошибка ввода-вывода: '
  'Msg.ErrUnexpected' = 'Непредвиденная ошибка: {0}'
  'Dlg.Title' = 'Подтверждение удаления'
  'Dlg.Heading' = 'Будет удалено элементов: {0}'
  'Dlg.Perm' = 'Режим: безвозвратное удаление. Файлы не попадут в корзину и не могут быть восстановлены.'
  'Dlg.Recycle' = 'Режим: перемещение в корзину. Элементы можно восстановить.'
  'Dlg.Confirm' = 'Я подтверждаю удаление перечисленных элементов'
  'Dlg.Cancel' = 'Отмена'
  'Dlg.Ok' = 'Удалить'
  'Export.Title' = 'Экспорт результатов поиска'
  'Export.FilterTxt' = 'Текстовый файл (*.txt)|*.txt'
  'Export.FileName' = 'Список_ZW3D_'
  'Mode.Permanent' = 'безвозвратное удаление'
  'Mode.Recycle' = 'в корзину'
  'ZwMode.Off' = 'выкл.'
  'ZwMode.All' = 'все файлы ZW3D'
  'ZwMode.Junk' = 'только временные ZW3D'
  'ZwMode.Model' = 'только модели ZW3D'
  'Dlg.UfTitle' = 'Очистка папки пользователя ZW3D'
  'Dlg.UfVersion' = 'Версия для очистки:'
  'Dlg.UfLocation' = 'Расположение:'
  'Dlg.UfHint' = 'Установленные версии ZW3D определяются автоматически; при смене версии пути ниже обновляются.'
  'Dlg.UfOpen' = 'Открыть'
  'Dlg.UfClose' = 'Закрыть'
  'Dlg.UfNote' = 'Внимание: очистка влияет на запущенный ZW3D — сначала закройте его. Перед удалением будет запрошено подтверждение.'
  'Uf.StatLine' = 'Файлов: {0} / папок: {1} / {2}'
  'Uf.NoTarget' = '(папка не найдена, очищать нечего)'
  'Uf.ClearOutput' = 'Очистить вывод'
  'Uf.ClearCustom' = 'Очистить настройки'
  'Uf.ClearTemp' = 'Очистить временные'
  'Uf.ClearAll' = 'Очистить папку польз.'
  'Uf.DescOutput' = 'Удаляет output\ (журналы, архивы, конвертированные файлы)'
  'Uf.DescCustom' = 'Удаляет custom\ (настройки интерфейса, шаблоны, лицензии — сбрасывает личные настройки)'
  'Uf.DescTemp' = 'Удаляет temp\ (файлы сеансов и промежуточные файлы)'
  'Uf.DescAll' = 'Удаляет всю папку версии (всё перечисленное выше и прочее)'
  'Uf.Confirm' = 'Действие: {0}¶Версия: {5}¶Папка: {1}¶Файлов: {2}, подпапок: {3}.¶Режим: {4}'
  'Uf.WarnCustom' = '¶¶Внимание: личные настройки будут удалены; ZW3D создаст их заново при следующем запуске.'
  'Uf.WarnAll' = '¶¶Внимание: вся папка пользователя будет очищена. Сначала сохраните нужные настройки.'
  'Uf.ZwRunning' = 'ZW3D запущен; файлы могут быть заблокированы. Рекомендуется сначала закрыть ZW3D. Продолжить?'
  'Uf.NoFolder' = 'Папка пользователя ZW3D не найдена: {0}'
  'Uf.Missing' = 'Папка не существует, очищать нечего: {0}'
  'Uf.Ok' = 'Очистка завершена: {0}'
  'Uf.Fail' = 'Ошибка очистки: {0}'
  'Uf.Cancelled' = 'Отменено'
  'Uf.ConfirmTitle' = 'Подтверждение очистки'
 }
 'pl_PL' = @{
  'Lang.All' = 'Wszystkie języki'
  'Hint' = 'Wskazówka: puste słowo kluczowe = pokaż wszystko (maks. 5000). Enter — szukaj.'
  'Zw3dHint' = 'Kryteria: lista rozszerzeń / nazwa zawiera ZW3D·TopologicRep·Proxy / ścieżka zawiera ZW3D'
  'ZwScope.All' = 'Wszystkie pliki ZW3D (modele + tymczasowe)'
  'ZwScope.Junk' = 'Tylko pliki tymczasowe ZW3D (clear ALL.bat)'
  'ZwScope.Model' = 'Tylko modele ZW3D (.z3/.z3prt/.z3asm/.z3drw)'
  'CreditTip' = 'Twórca'
  'Tip.Search' = 'Rozpocznij wyszukiwanie (lub naciśnij Enter w polu słowa kluczowego)'
  'Tip.Stop' = 'Zatrzymaj bieżące skanowanie'
  'Tip.CheckAll' = 'Zaznacz wszystkie wyniki (do usunięcia)'
  'Tip.Uncheck' = 'Usuń wszystkie zaznaczenia'
  'Tip.Delete' = 'Usuń zaznaczone elementy. Domyślnie trafiają do kosza (można przywrócić); zaznacz „Usuń trwale”, aby usunąć je od razu.'
  'Tip.Explorer' = 'Otwórz zaznaczony element w Eksploratorze'
  'Tip.LogFolder' = 'Otwórz folder dzienników (logs\)'
  'Tip.Export' = 'Zapisz wyniki jako .txt / .csv / .json (format zgodny z polem rozszerzenia)'
  'Tip.Abort' = 'Przerwij wszystkie operacje i zakończ (Ctrl+Q). Bez dodatkowego pytania; szczegóły w dzienniku.'
  'Tip.Refresh' = 'Odśwież listę dysków (po podłączeniu USB)'
  'Tip.ToggleDrives' = 'Przełącz zaznaczenie wszystkich dysków'
  'Tip.UiLang' = 'Zmienia tylko język interfejsu tego narzędzia; nie wpływa na ZW3D ani na wyniki. Wybór jest zapamiętywany.'
  'Tip.ScanLang' = 'Wersja językowa ZW3D używana podczas skanowania: wpływa na szablony słów kluczowych i zlokalizowane terminy. „Wszystkie języki” = bez ograniczeń. Nie zmienia interfejsu.'
  'Tip.UserFolder' = 'Wyczyść folder użytkownika ZW3D tymi samymi czterema akcjami co oficjalny User Folder Manager'
  'Drive.Available' = '({0} / {1} GB wolne)'
  'Status.ReadyFull' = 'Gotowe · wybierz dysk i wpisz słowo kluczowe'
  'Status.Scanning' = 'Skanowanie…'
  'Status.ScanProgress' = 'Skanowanie… sprawdzono {0}, trafień {1}, pominięto {2} niedostępnych folderów'
  'Status.Found' = 'Znaleziono {0} (przeskanowano {1}, pominięto {2} niedostępnych folderów)'
  'Status.Truncated' = ' · Osiągnięto limit 5000 wyników, zawęź zakres'
  'Status.ExtHint' = ' · Wskazówka: {0} nie ma na liście filtra ZW3D; odznacz „Tylko pliki ZW3D”'
  'Status.Checked' = 'Zaznaczono: {0}'
  'Status.Unchecked' = 'Zaznaczenie wyczyszczone'
  'Status.DrivesRefreshed' = 'Lista dysków odświeżona'
  'Status.Exported' = 'Wyeksportowano wierszy: {0} -> {1}'
  'Status.NoDelete' = 'Nic do usunięcia'
  'Status.Cancelled' = 'Usuwanie anulowane'
  'Status.SearchFail' = 'Nie udało się uruchomić wyszukiwania, zobacz dziennik'
  'Status.Aborting' = 'Zatrzymywanie: przerywanie wszystkich operacji…'
  'Status.Deleting' = 'Usuwanie {0}/{1}: {2}'
  'Status.Summary' = 'Powodzenie: {0}, błędy: {1} ({2})'
  'Msg.InfoTitle' = 'Informacja'
  'Msg.ErrTitle' = 'Błąd'
  'Msg.WarnTitle' = 'Ostrzeżenie'
  'Msg.NoDrive' = 'Zaznacz co najmniej jeden dysk.'
  'Msg.NoExport' = 'Brak wyników do eksportu. Najpierw wykonaj wyszukiwanie.'
  'Msg.NoSelection' = 'Najpierw zaznacz elementy do usunięcia.'
  'Msg.ProtectedTitle' = 'Pominięto chronione ścieżki'
  'Msg.ProtectedBody' = 'Pominięto następujące chronione ścieżki systemowe:¶'
  'Msg.ExportOkTitle' = 'Eksport zakończony'
  'Msg.ExportOk' = 'Wyeksportowano wierszy: {0} do:¶{1}'
  'Msg.ExportFailBody' = 'Eksport nie powiódł się:¶'
  'Msg.PartialTitle' = 'Nie udało się usunąć niektórych elementów'
  'Msg.ResultTitle' = 'Gotowe'
  'Msg.ResultBody' = 'Usuwanie zakończone.¶¶{0}'
  'Msg.FailListBody' = 'Wynik: {0}¶¶Nieudane elementy:¶'
  'Msg.RecycleFail' = 'Składnik kosza jest niedostępny; zaznacz „Usuń trwale” i spróbuj ponownie'
  'Msg.ErrPermission' = 'Odmowa dostępu (UnauthorizedAccess)'
  'Msg.ErrSecurity' = 'Ograniczenie bezpieczeństwa (SecurityException)'
  'Msg.ErrLocked' = 'Plik / folder jest używany przez inny proces: '
  'Msg.ErrLongPath' = 'Ścieżka zbyt długa lub nieprawidłowa: '
  'Msg.ErrIO' = 'Błąd we/wy: '
  'Msg.ErrUnexpected' = 'Nieoczekiwany błąd: {0}'
  'Dlg.Title' = 'Potwierdź usunięcie'
  'Dlg.Heading' = 'Do usunięcia: {0} element(ów)'
  'Dlg.Perm' = 'Tryb: trwałe usunięcie. Pliki nie trafią do kosza i nie będzie można ich przywrócić.'
  'Dlg.Recycle' = 'Tryb: przenieś do kosza. Elementy można przywrócić.'
  'Dlg.Confirm' = 'Potwierdzam usunięcie powyższych elementów'
  'Dlg.Cancel' = 'Anuluj'
  'Dlg.Ok' = 'Usuń'
  'Export.Title' = 'Eksportuj wyniki wyszukiwania'
  'Export.FilterTxt' = 'Plik tekstowy (*.txt)|*.txt'
  'Export.FileName' = 'Lista_ZW3D_'
  'Mode.Permanent' = 'trwałe usunięcie'
  'Mode.Recycle' = 'przeniesiono do kosza'
  'ZwMode.Off' = 'wył.'
  'ZwMode.All' = 'wszystkie pliki ZW3D'
  'ZwMode.Junk' = 'tylko tymczasowe ZW3D'
  'ZwMode.Model' = 'tylko modele ZW3D'
  'Dlg.UfTitle' = 'Czyszczenie folderu użytkownika ZW3D'
  'Dlg.UfVersion' = 'Wersja do wyczyszczenia:'
  'Dlg.UfLocation' = 'Lokalizacja:'
  'Dlg.UfHint' = 'Zainstalowane wersje ZW3D są wykrywane automatycznie; po zmianie wersji ścieżki poniżej się aktualizują.'
  'Dlg.UfOpen' = 'Otwórz'
  'Dlg.UfClose' = 'Zamknij'
  'Dlg.UfNote' = 'Uwaga: czyszczenie wpływa na działający ZW3D — najpierw go zamknij. Przed usunięciem pojawi się potwierdzenie.'
  'Uf.StatLine' = 'Pliki: {0} / foldery: {1} / {2}'
  'Uf.NoTarget' = '(folder nie istnieje, nic do czyszczenia)'
  'Uf.ClearOutput' = 'Wyczyść pliki wyjściowe'
  'Uf.ClearCustom' = 'Wyczyść konfigurację'
  'Uf.ClearTemp' = 'Wyczyść pliki tymczasowe'
  'Uf.ClearAll' = 'Wyczyść folder użytk.'
  'Uf.DescOutput' = 'Usuwa output\ (dzienniki, archiwa, pliki skonwertowane)'
  'Uf.DescCustom' = 'Usuwa custom\ (konfiguracja UI, szablony, licencje — resetuje ustawienia osobiste)'
  'Uf.DescTemp' = 'Usuwa temp\ (pliki sesji i pośrednie)'
  'Uf.DescAll' = 'Usuwa cały folder wersji (wszystko powyżej i więcej)'
  'Uf.Confirm' = 'Akcja: {0}¶Wersja: {5}¶Folder: {1}¶Pliki: {2}, podfoldery: {3}.¶Tryb: {4}'
  'Uf.WarnCustom' = '¶¶Ostrzeżenie: konfiguracje osobiste zostaną usunięte; ZW3D odtworzy je przy następnym uruchomieniu.'
  'Uf.WarnAll' = '¶¶Ostrzeżenie: cały folder użytkownika zostanie opróżniony. Najpierw zrób kopię potrzebnych ustawień.'
  'Uf.ZwRunning' = 'ZW3D jest uruchomiony; pliki mogą być zablokowane. Zalecane jest najpierw zamknięcie ZW3D. Kontynuować mimo to?'
  'Uf.NoFolder' = 'Nie znaleziono folderu użytkownika ZW3D: {0}'
  'Uf.Missing' = 'Folder nie istnieje, nic do czyszczenia: {0}'
  'Uf.Ok' = 'Czyszczenie zakończone: {0}'
  'Uf.Fail' = 'Czyszczenie nie powiodło się: {0}'
  'Uf.Cancelled' = 'Anulowano'
  'Uf.ConfirmTitle' = 'Potwierdź czyszczenie'
 }
 'cs_CZ' = @{
  'Lang.All' = 'Všechny jazyky'
  'Hint' = 'Tip: prázdné klíčové slovo = vypsat vše (max. 5000 výsledků). Enter spustí hledání.'
  'Zw3dHint' = 'Kritéria: seznam přípon / název obsahuje ZW3D·TopologicRep·Proxy / cesta obsahuje ZW3D'
  'ZwScope.All' = 'Všechny soubory ZW3D (modely + dočasné)'
  'ZwScope.Junk' = 'Pouze dočasné soubory ZW3D (clear ALL.bat)'
  'ZwScope.Model' = 'Pouze modely ZW3D (.z3/.z3prt/.z3asm/.z3drw)'
  'CreditTip' = 'Vývojář'
  'Tip.Search' = 'Spustit hledání (nebo stiskněte Enter v poli klíčového slova)'
  'Tip.Stop' = 'Zastavit probíhající prohledávání'
  'Tip.CheckAll' = 'Zaškrtnout všechny výsledky (připravit ke smazání)'
  'Tip.Uncheck' = 'Zrušit všechna zaškrtnutí'
  'Tip.Delete' = 'Smazat zaškrtnuté položky. Ve výchozím stavu jdou do koše (lze obnovit); zaškrtněte „Trvale smazat“ pro okamžité odstranění.'
  'Tip.Explorer' = 'Otevřít vybranou položku v Průzkumníku'
  'Tip.LogFolder' = 'Otevřít složku protokolů (logs\)'
  'Tip.Export' = 'Uložit výsledky jako .txt / .csv / .json (formát podle pole přípony)'
  'Tip.Abort' = 'Přerušit všechny operace a ukončit (Ctrl+Q). Bez dalšího dotazu; podrobnosti jsou v protokolu.'
  'Tip.Refresh' = 'Znovu načíst seznam disků (po připojení USB)'
  'Tip.ToggleDrives' = 'Přepnout zaškrtnutí všech disků'
  'Tip.UiLang' = 'Mění pouze jazyk rozhraní tohoto nástroje; neovlivní ZW3D ani výsledky hledání. Volba se zapamatuje.'
  'Tip.ScanLang' = 'Jazyková verze ZW3D pro prohledávání: ovlivňuje šablony klíčových slov a lokalizované výrazy. „Všechny jazyky“ = bez omezení. Rozhraní se nemění.'
  'Tip.UserFolder' = 'Vyčistit uživatelskou složku ZW3D stejnými čtyřmi akcemi jako oficiální User Folder Manager'
  'Drive.Available' = '({0} / {1} GB volno)'
  'Status.ReadyFull' = 'Připraveno · vyberte disk a zadejte klíčové slovo'
  'Status.Scanning' = 'Prohledávání…'
  'Status.ScanProgress' = 'Prohledávání… zkontrolováno {0}, nalezeno {1}, přeskočeno {2} nepřístupných složek'
  'Status.Found' = 'Nalezeno {0} (prohledáno {1}, přeskočeno {2} nepřístupných složek)'
  'Status.Truncated' = ' · Dosažen limit 5000 výsledků, zužte rozsah'
  'Status.ExtHint' = ' · Tip: {0} není v seznamu filtru ZW3D; zrušte „Pouze soubory ZW3D“'
  'Status.Checked' = 'Zaškrtnuto: {0}'
  'Status.Unchecked' = 'Výběr zrušen'
  'Status.DrivesRefreshed' = 'Seznam disků obnoven'
  'Status.Exported' = 'Exportováno řádků: {0} -> {1}'
  'Status.NoDelete' = 'Není co smazat'
  'Status.Cancelled' = 'Mazání zrušeno'
  'Status.SearchFail' = 'Hledání nelze spustit, viz protokol'
  'Status.Aborting' = 'Ukončování: přerušují se všechny operace…'
  'Status.Deleting' = 'Mazání {0}/{1}: {2}'
  'Status.Summary' = 'Úspěšně: {0}, chyb: {1} ({2})'
  'Msg.InfoTitle' = 'Informace'
  'Msg.ErrTitle' = 'Chyba'
  'Msg.WarnTitle' = 'Upozornění'
  'Msg.NoDrive' = 'Vyberte alespoň jeden disk.'
  'Msg.NoExport' = 'Nejsou žádné výsledky k exportu. Nejprve spusťte hledání.'
  'Msg.NoSelection' = 'Nejprve zaškrtněte položky ke smazání.'
  'Msg.ProtectedTitle' = 'Chráněné cesty přeskočeny'
  'Msg.ProtectedBody' = 'Následující chráněné systémové cesty byly přeskočeny:¶'
  'Msg.ExportOkTitle' = 'Export dokončen'
  'Msg.ExportOk' = 'Exportováno řádků: {0} do:¶{1}'
  'Msg.ExportFailBody' = 'Export se nezdařil:¶'
  'Msg.PartialTitle' = 'Některé položky se nepodařilo smazat'
  'Msg.ResultTitle' = 'Hotovo'
  'Msg.ResultBody' = 'Mazání dokončeno.¶¶{0}'
  'Msg.FailListBody' = 'Výsledek: {0}¶¶Neúspěšné položky:¶'
  'Msg.RecycleFail' = 'Komponenta koše není dostupná; zaškrtněte „Trvale smazat“ a zkuste to znovu'
  'Msg.ErrPermission' = 'Přístup odepřen (UnauthorizedAccess)'
  'Msg.ErrSecurity' = 'Bezpečnostní omezení (SecurityException)'
  'Msg.ErrLocked' = 'Soubor / složku používá jiný proces: '
  'Msg.ErrLongPath' = 'Cesta je příliš dlouhá nebo neplatná: '
  'Msg.ErrIO' = 'Chyba V/V: '
  'Msg.ErrUnexpected' = 'Neočekávaná chyba: {0}'
  'Dlg.Title' = 'Potvrdit smazání'
  'Dlg.Heading' = 'Bude smazáno položek: {0}'
  'Dlg.Perm' = 'Režim: trvalé smazání. Soubory neprojdou košem a nelze je obnovit.'
  'Dlg.Recycle' = 'Režim: přesunout do koše. Položky lze obnovit.'
  'Dlg.Confirm' = 'Potvrzuji smazání výše uvedených položek'
  'Dlg.Cancel' = 'Zrušit'
  'Dlg.Ok' = 'Smazat'
  'Export.Title' = 'Exportovat výsledky hledání'
  'Export.FilterTxt' = 'Textový soubor (*.txt)|*.txt'
  'Export.FileName' = 'Seznam_ZW3D_'
  'Mode.Permanent' = 'trvalé smazání'
  'Mode.Recycle' = 'přesunuto do koše'
  'ZwMode.Off' = 'vypnuto'
  'ZwMode.All' = 'všechny soubory ZW3D'
  'ZwMode.Junk' = 'pouze dočasné ZW3D'
  'ZwMode.Model' = 'pouze modely ZW3D'
  'Dlg.UfTitle' = 'Čištění uživatelské složky ZW3D'
  'Dlg.UfVersion' = 'Verze k vyčištění:'
  'Dlg.UfLocation' = 'Umístění:'
  'Dlg.UfHint' = 'Nainstalované verze ZW3D se zjistí automaticky; po změně verze se cesty níže aktualizují.'
  'Dlg.UfOpen' = 'Otevřít'
  'Dlg.UfClose' = 'Zavřít'
  'Dlg.UfNote' = 'Pozor: čištění ovlivní spuštěný ZW3D — nejprve jej zavřete. Před smazáním budete požádáni o potvrzení.'
  'Uf.StatLine' = 'Soubory: {0} / složky: {1} / {2}'
  'Uf.NoTarget' = '(složka neexistuje, není co čistit)'
  'Uf.ClearOutput' = 'Vymazat výstupní soubory'
  'Uf.ClearCustom' = 'Vymazat konfiguraci'
  'Uf.ClearTemp' = 'Vymazat dočasné soubory'
  'Uf.ClearAll' = 'Vyprázdnit složku uživ.'
  'Uf.DescOutput' = 'Smaže output\ (protokoly, archivy, převedené soubory)'
  'Uf.DescCustom' = 'Smaže custom\ (konfigurace UI, šablony, licence — obnoví osobní nastavení)'
  'Uf.DescTemp' = 'Smaže temp\ (soubory relací a mezivýsledky)'
  'Uf.DescAll' = 'Smaže celou složku verze (vše výše uvedené a další)'
  'Uf.Confirm' = 'Akce: {0}¶Verze: {5}¶Složka: {1}¶Soubory: {2}, podsložky: {3}.¶Režim: {4}'
  'Uf.WarnCustom' = '¶¶Upozornění: osobní konfigurace budou odstraněny; ZW3D je při dalším spuštění vytvoří znovu.'
  'Uf.WarnAll' = '¶¶Upozornění: celá uživatelská složka bude vyprázdněna. Nejprve zálohujte potřebná nastavení.'
  'Uf.ZwRunning' = 'ZW3D je spuštěn; soubory mohou být uzamčeny. Doporučujeme ZW3D nejprve zavřít. Přesto pokračovat?'
  'Uf.NoFolder' = 'Uživatelská složka ZW3D nenalezena: {0}'
  'Uf.Missing' = 'Složka neexistuje, není co čistit: {0}'
  'Uf.Ok' = 'Čištění dokončeno: {0}'
  'Uf.Fail' = 'Čištění se nezdařilo: {0}'
  'Uf.Cancelled' = 'Zrušeno'
  'Uf.ConfirmTitle' = 'Potvrdit čištění'
 }
 'hu_HU' = @{
  'Lang.All' = 'Minden nyelv'
  'Hint' = 'Tipp: üres kulcsszó = minden listázása (max. 5000 találat). Enter: keresés.'
  'Zw3dHint' = 'Feltétel: kiterjesztéslista / a név tartalmazza: ZW3D·TopologicRep·Proxy / az útvonal: ZW3D'
  'ZwScope.All' = 'Minden ZW3D fájl (modellek + ideiglenes)'
  'ZwScope.Junk' = 'Csak ZW3D ideiglenes fájlok (clear ALL.bat)'
  'ZwScope.Model' = 'Csak ZW3D modellek (.z3/.z3prt/.z3asm/.z3drw)'
  'CreditTip' = 'Fejlesztő'
  'Tip.Search' = 'Keresés indítása (vagy Enter a kulcsszó mezőben)'
  'Tip.Stop' = 'A folyamatban lévő keresés leállítása'
  'Tip.CheckAll' = 'Minden találat bejelölése (törlésre)'
  'Tip.Uncheck' = 'Minden jelölés törlése'
  'Tip.Delete' = 'A bejelölt elemek törlése. Alapértelmezés szerint a Lomtárba kerülnek (visszaállítható); a „Végleges törlés” bejelölésével azonnal törlődnek.'
  'Tip.Explorer' = 'A kijelölt elem megnyitása az Intézőben'
  'Tip.LogFolder' = 'A naplómappa megnyitása (logs\)'
  'Tip.Export' = 'Találatok mentése .txt / .csv / .json formátumban (a formátum a kiterjesztés mezőt követi)'
  'Tip.Abort' = 'Minden művelet megszakítása és kilépés (Ctrl+Q). Nincs további kérdés; részletek a naplóban.'
  'Tip.Refresh' = 'Meghajtólista frissítése (USB csatlakoztatása után)'
  'Tip.ToggleDrives' = 'Az összes meghajtó jelölésének átváltása'
  'Tip.UiLang' = 'Csak az eszköz felületének nyelvét váltja; a ZW3D-t és a találatokat nem érinti. A választás megmarad.'
  'Tip.ScanLang' = 'A kereséskor használt ZW3D nyelvi verzió: hat a kulcsszósablonokra és a honosított kifejezésekre. „Minden nyelv” = nincs korlátozás. A felületet nem változtatja.'
  'Tip.UserFolder' = 'A ZW3D felhasználói mappa tisztítása a hivatalos User Folder Manager négy műveletével'
  'Drive.Available' = '({0} / {1} GB szabad)'
  'Status.ReadyFull' = 'Kész · válasszon meghajtót és adjon meg kulcsszót'
  'Status.Scanning' = 'Keresés…'
  'Status.ScanProgress' = 'Keresés… {0} ellenőrizve, {1} találat, {2} elérhetetlen mappa kihagyva'
  'Status.Found' = '{0} találat ({1} ellenőrizve, {2} elérhetetlen mappa kihagyva)'
  'Status.Truncated' = ' · Elérte az 5000 találatos korlátot, szűkítse a keresést'
  'Status.ExtHint' = ' · Tipp: {0} nincs a ZW3D szűrőlistában; vegye ki a „Csak ZW3D fájlok” jelölést'
  'Status.Checked' = '{0} elem bejelölve'
  'Status.Unchecked' = 'Jelölések törölve'
  'Status.DrivesRefreshed' = 'Meghajtólista frissítve'
  'Status.Exported' = '{0} sor exportálva -> {1}'
  'Status.NoDelete' = 'Nincs mit törölni'
  'Status.Cancelled' = 'Törlés megszakítva'
  'Status.SearchFail' = 'A keresés nem indítható, lásd a naplót'
  'Status.Aborting' = 'Leállítás: minden művelet megszakítása…'
  'Status.Deleting' = 'Törlés {0}/{1}: {2}'
  'Status.Summary' = '{0} sikeres, {1} sikertelen ({2})'
  'Msg.InfoTitle' = 'Információ'
  'Msg.ErrTitle' = 'Hiba'
  'Msg.WarnTitle' = 'Figyelmeztetés'
  'Msg.NoDrive' = 'Jelöljön ki legalább egy meghajtót.'
  'Msg.NoExport' = 'Nincs exportálható találat. Előbb végezzen keresést.'
  'Msg.NoSelection' = 'Előbb jelölje be a törlendő elemeket.'
  'Msg.ProtectedTitle' = 'Védett útvonalak kihagyva'
  'Msg.ProtectedBody' = 'A következő védett rendszerútvonalak ki lettek hagyva:¶'
  'Msg.ExportOkTitle' = 'Exportálás kész'
  'Msg.ExportOk' = '{0} sor exportálva ide:¶{1}'
  'Msg.ExportFailBody' = 'Az exportálás sikertelen:¶'
  'Msg.PartialTitle' = 'Néhány elemet nem sikerült törölni'
  'Msg.ResultTitle' = 'Kész'
  'Msg.ResultBody' = 'A törlés befejeződött.¶¶{0}'
  'Msg.FailListBody' = 'Eredmény: {0}¶¶Sikertelen elemek:¶'
  'Msg.RecycleFail' = 'A Lomtár összetevő nem érhető el; jelölje be a „Végleges törlés” opciót, és próbálja újra'
  'Msg.ErrPermission' = 'Hozzáférés megtagadva (UnauthorizedAccess)'
  'Msg.ErrSecurity' = 'Biztonsági korlátozás (SecurityException)'
  'Msg.ErrLocked' = 'A fájlt / mappát egy másik folyamat használja: '
  'Msg.ErrLongPath' = 'Az útvonal túl hosszú vagy érvénytelen: '
  'Msg.ErrIO' = 'I/O hiba: '
  'Msg.ErrUnexpected' = 'Váratlan hiba: {0}'
  'Dlg.Title' = 'Törlés megerősítése'
  'Dlg.Heading' = '{0} elem törlésre kerül'
  'Dlg.Perm' = 'Mód: végleges törlés. A fájlok nem kerülnek a Lomtárba, és nem állíthatók vissza.'
  'Dlg.Recycle' = 'Mód: áthelyezés a Lomtárba. Az elemek visszaállíthatók.'
  'Dlg.Confirm' = 'Megerősítem a fenti elemek törlését'
  'Dlg.Cancel' = 'Mégse'
  'Dlg.Ok' = 'Törlés'
  'Export.Title' = 'Találatok exportálása'
  'Export.FilterTxt' = 'Szövegfájl (*.txt)|*.txt'
  'Export.FileName' = 'ZW3D_lista_'
  'Mode.Permanent' = 'végleges törlés'
  'Mode.Recycle' = 'Lomtárba helyezve'
  'ZwMode.Off' = 'ki'
  'ZwMode.All' = 'minden ZW3D fájl'
  'ZwMode.Junk' = 'csak ZW3D ideiglenes'
  'ZwMode.Model' = 'csak ZW3D modellek'
  'Dlg.UfTitle' = 'ZW3D felhasználói mappa tisztítása'
  'Dlg.UfVersion' = 'Tisztítandó verzió:'
  'Dlg.UfLocation' = 'Hely:'
  'Dlg.UfHint' = 'A telepített ZW3D verziók felismerése automatikus; verzióváltáskor az alábbi útvonalak frissülnek.'
  'Dlg.UfOpen' = 'Megnyitás'
  'Dlg.UfClose' = 'Bezárás'
  'Dlg.UfNote' = 'Figyelem: a tisztítás hat a futó ZW3D-re — előbb zárja be. Törlés előtt megerősítést kérünk.'
  'Uf.StatLine' = '{0} fájl / {1} mappa / {2}'
  'Uf.NoTarget' = '(a mappa nem létezik, nincs mit tisztítani)'
  'Uf.ClearOutput' = 'Kimeneti fájlok törlése'
  'Uf.ClearCustom' = 'Egyéni beállítások törlése'
  'Uf.ClearTemp' = 'Ideiglenes fájlok törlése'
  'Uf.ClearAll' = 'Felhasználói mappa ürítése'
  'Uf.DescOutput' = 'Törli: output\ (naplók, archívumok, konvertált fájlok)'
  'Uf.DescCustom' = 'Törli: custom\ (felület-beállítások, sablonok, licencek — visszaállítja a személyes beállításokat)'
  'Uf.DescTemp' = 'Törli: temp\ (munkamenet- és köztes fájlok)'
  'Uf.DescAll' = 'Törli a teljes verziómappát (minden fentit és egyebet)'
  'Uf.Confirm' = 'Művelet: {0}¶Verzió: {5}¶Mappa: {1}¶{2} fájl, {3} almappa.¶Mód: {4}'
  'Uf.WarnCustom' = '¶¶Figyelem: a személyes beállítások törlődnek; a ZW3D a következő indításkor újra létrehozza őket.'
  'Uf.WarnAll' = '¶¶Figyelem: a teljes felhasználói mappa kiürül. Előbb mentse a szükséges beállításokat.'
  'Uf.ZwRunning' = 'A ZW3D fut; a fájlok zárolva lehetnek. Javasolt előbb bezárni a ZW3D-t. Mégis folytatja?'
  'Uf.NoFolder' = 'A ZW3D felhasználói mappa nem található: {0}'
  'Uf.Missing' = 'A mappa nem létezik, nincs mit tisztítani: {0}'
  'Uf.Ok' = 'Tisztítás kész: {0}'
  'Uf.Fail' = 'A tisztítás sikertelen: {0}'
  'Uf.Cancelled' = 'Megszakítva'
  'Uf.ConfirmTitle' = 'Tisztítás megerősítése'
 }
 'tr_TR' = @{
  'Lang.All' = 'Tüm diller'
  'Hint' = 'İpucu: anahtar kelime boş = tümünü listele (en fazla 5000). Aramak için Enter.'
  'Zw3dHint' = 'Ölçüt: uzantı listesi / adda ZW3D·TopologicRep·Proxy / yolda ZW3D'
  'ZwScope.All' = 'Tüm ZW3D dosyaları (modeller + geçici)'
  'ZwScope.Junk' = 'Yalnızca ZW3D geçici dosyaları (clear ALL.bat)'
  'ZwScope.Model' = 'Yalnızca ZW3D modelleri (.z3/.z3prt/.z3asm/.z3drw)'
  'CreditTip' = 'Geliştirici'
  'Tip.Search' = 'Aramayı başlat (veya anahtar kelime kutusunda Enter''a basın)'
  'Tip.Stop' = 'Devam eden taramayı durdur'
  'Tip.CheckAll' = 'Tüm sonuçları işaretle (silmeye hazır)'
  'Tip.Uncheck' = 'Tüm işaretleri kaldır'
  'Tip.Delete' = 'İşaretli öğeleri sil. Varsayılan olarak Geri Dönüşüm Kutusu''na gider (geri yüklenebilir); doğrudan silmek için “Kalıcı sil”i işaretleyin.'
  'Tip.Explorer' = 'Seçili öğeyi Dosya Gezgini''nde aç'
  'Tip.LogFolder' = 'Günlük klasörünü aç (logs\)'
  'Tip.Export' = 'Sonuçları .txt / .csv / .json olarak kaydet (biçim uzantı kutusuna göre)'
  'Tip.Abort' = 'Tüm işlemleri iptal edip çık (Ctrl+Q). Ek onay yok; ayrıntılar günlüğe yazılır.'
  'Tip.Refresh' = 'Sürücü listesini yenile (USB taktıktan sonra)'
  'Tip.ToggleDrives' = 'Tüm sürücülerin işaretini değiştir'
  'Tip.UiLang' = 'Yalnızca bu aracın arayüz dilini değiştirir; ZW3D''yi veya arama sonuçlarını etkilemez. Seçim hatırlanır.'
  'Tip.ScanLang' = 'Tarama sırasında kullanılan ZW3D dil sürümü: hazır anahtar kelimeleri ve yerelleştirilmiş terimleri etkiler. “Tüm diller” = kısıtlama yok. Arayüzü değiştirmez.'
  'Tip.UserFolder' = 'ZW3D kullanıcı klasörünü resmi User Folder Manager ile aynı dört işlemle temizle'
  'Drive.Available' = '({0} / {1} GB boş)'
  'Status.ReadyFull' = 'Hazır · bir sürücü seçip anahtar kelime girin'
  'Status.Scanning' = 'Taranıyor…'
  'Status.ScanProgress' = 'Taranıyor… {0} denetlendi, {1} eşleşme, erişilemeyen {2} klasör atlandı'
  'Status.Found' = '{0} bulundu ({1} tarandı, erişilemeyen {2} klasör atlandı)'
  'Status.Truncated' = ' · 5000 sonuç sınırına ulaşıldı, kapsamı daraltın'
  'Status.ExtHint' = ' · İpucu: {0} ZW3D filtre listesinde yok; “Yalnızca ZW3D dosyaları” işaretini kaldırın'
  'Status.Checked' = '{0} öğe işaretlendi'
  'Status.Unchecked' = 'İşaretler kaldırıldı'
  'Status.DrivesRefreshed' = 'Sürücü listesi yenilendi'
  'Status.Exported' = '{0} satır dışa aktarıldı -> {1}'
  'Status.NoDelete' = 'Silinecek öğe yok'
  'Status.Cancelled' = 'Silme iptal edildi'
  'Status.SearchFail' = 'Arama başlatılamadı, günlüğe bakın'
  'Status.Aborting' = 'Durduruluyor: tüm işlemler iptal ediliyor…'
  'Status.Deleting' = 'Siliniyor {0}/{1}: {2}'
  'Status.Summary' = '{0} başarılı, {1} başarısız ({2})'
  'Msg.InfoTitle' = 'Bilgi'
  'Msg.ErrTitle' = 'Hata'
  'Msg.WarnTitle' = 'Uyarı'
  'Msg.NoDrive' = 'En az bir sürücü seçin.'
  'Msg.NoExport' = 'Dışa aktarılacak sonuç yok. Önce arama yapın.'
  'Msg.NoSelection' = 'Önce silinecek öğeleri işaretleyin.'
  'Msg.ProtectedTitle' = 'Korumalı yollar atlandı'
  'Msg.ProtectedBody' = 'Aşağıdaki korumalı sistem yolları atlandı:¶'
  'Msg.ExportOkTitle' = 'Dışa aktarma tamamlandı'
  'Msg.ExportOk' = '{0} satır şuraya aktarıldı:¶{1}'
  'Msg.ExportFailBody' = 'Dışa aktarma başarısız:¶'
  'Msg.PartialTitle' = 'Bazı öğeler silinemedi'
  'Msg.ResultTitle' = 'Tamamlandı'
  'Msg.ResultBody' = 'Silme tamamlandı.¶¶{0}'
  'Msg.FailListBody' = 'Sonuç: {0}¶¶Başarısız öğeler:¶'
  'Msg.RecycleFail' = 'Geri Dönüşüm Kutusu bileşeni kullanılamıyor; “Kalıcı sil”i işaretleyip yeniden deneyin'
  'Msg.ErrPermission' = 'Erişim reddedildi (UnauthorizedAccess)'
  'Msg.ErrSecurity' = 'Güvenlik kısıtlaması (SecurityException)'
  'Msg.ErrLocked' = 'Dosya / klasör başka bir işlem tarafından kullanılıyor: '
  'Msg.ErrLongPath' = 'Yol çok uzun veya geçersiz: '
  'Msg.ErrIO' = 'G/Ç hatası: '
  'Msg.ErrUnexpected' = 'Beklenmeyen hata: {0}'
  'Dlg.Title' = 'Silmeyi onayla'
  'Dlg.Heading' = '{0} öğe silinecek'
  'Dlg.Perm' = 'Mod: kalıcı silme. Dosyalar Geri Dönüşüm Kutusu''na gitmez ve geri yüklenemez.'
  'Dlg.Recycle' = 'Mod: Geri Dönüşüm Kutusu''na taşı. Öğeler geri yüklenebilir.'
  'Dlg.Confirm' = 'Yukarıdaki öğeleri silmeyi onaylıyorum'
  'Dlg.Cancel' = 'İptal'
  'Dlg.Ok' = 'Sil'
  'Export.Title' = 'Arama sonuçlarını dışa aktar'
  'Export.FilterTxt' = 'Metin dosyası (*.txt)|*.txt'
  'Export.FileName' = 'ZW3D_liste_'
  'Mode.Permanent' = 'kalıcı silme'
  'Mode.Recycle' = 'Geri Dönüşüm Kutusu''na taşındı'
  'ZwMode.Off' = 'kapalı'
  'ZwMode.All' = 'tüm ZW3D dosyaları'
  'ZwMode.Junk' = 'yalnızca ZW3D geçici'
  'ZwMode.Model' = 'yalnızca ZW3D modelleri'
  'Dlg.UfTitle' = 'ZW3D kullanıcı klasörü temizliği'
  'Dlg.UfVersion' = 'Temizlenecek sürüm:'
  'Dlg.UfLocation' = 'Konum:'
  'Dlg.UfHint' = 'Kurulu ZW3D sürümleri otomatik algılanır; sürüm değiştirildiğinde aşağıdaki yollar güncellenir.'
  'Dlg.UfOpen' = 'Aç'
  'Dlg.UfClose' = 'Kapat'
  'Dlg.UfNote' = 'Not: temizlik çalışan ZW3D''yi etkiler; önce kapatın. Silmeden önce onay istenir.'
  'Uf.StatLine' = '{0} dosya / {1} klasör / {2}'
  'Uf.NoTarget' = '(klasör yok, temizlenecek bir şey yok)'
  'Uf.ClearOutput' = 'Çıktı dosyalarını temizle'
  'Uf.ClearCustom' = 'Özel yapılandırmayı temizle'
  'Uf.ClearTemp' = 'Geçici dosyaları temizle'
  'Uf.ClearAll' = 'Kullanıcı klasörünü boşalt'
  'Uf.DescOutput' = 'output\ silinir (günlükler, arşivler, dönüştürülmüş dosyalar)'
  'Uf.DescCustom' = 'custom\ silinir (arayüz yapılandırması, şablonlar, lisanslar — kişisel ayarlar sıfırlanır)'
  'Uf.DescTemp' = 'temp\ silinir (oturum ve ara dosyalar)'
  'Uf.DescAll' = 'Tüm sürüm klasörü silinir (yukarıdakilerin tümü ve fazlası)'
  'Uf.Confirm' = 'İşlem: {0}¶Sürüm: {5}¶Klasör: {1}¶{2} dosya, {3} alt klasör.¶Mod: {4}'
  'Uf.WarnCustom' = '¶¶Uyarı: kişisel yapılandırmalar kaldırılacak; ZW3D bir sonraki başlatmada bunları yeniden oluşturur.'
  'Uf.WarnAll' = '¶¶Uyarı: tüm kullanıcı klasörü boşaltılacak. Gerekli ayarları önce yedekleyin.'
  'Uf.ZwRunning' = 'ZW3D çalışıyor; dosyalar kilitli olabilir. Önce ZW3D''yi kapatmanız önerilir. Yine de devam edilsin mi?'
  'Uf.NoFolder' = 'ZW3D kullanıcı klasörü bulunamadı: {0}'
  'Uf.Missing' = 'Klasör yok, temizlenecek bir şey yok: {0}'
  'Uf.Ok' = 'Temizlik tamamlandı: {0}'
  'Uf.Fail' = 'Temizlik başarısız: {0}'
  'Uf.Cancelled' = 'İptal edildi'
  'Uf.ConfirmTitle' = 'Temizliği onayla'
 }
}
foreach ($lang in $script:StrMore.Keys) {
    $tbl = $script:StrMore[$lang]
    foreach ($k in $tbl.Keys) {
        if (-not $script:Str.ContainsKey($k)) { $script:Str[$k] = @{} }
        $script:Str[$k][$lang] = ([string]$tbl[$k]).Replace('¶', "`r`n")
    }
}
$script:StrMore = $null

# 取詞條：目前語言 → 英文 → 繁中（原版）
function L([string]$Key) {
    $row = $script:Str[$Key]
    if (-not $row) { return $Key }
    $v = $row[$script:UiLang]
    if ($v) { return [string]$v }
    $v = $row['en_US']
    if ($v) { return [string]$v }
    return [string]$row['zh_TW']
}

# 各語言版本的使用者檔案命名習慣（備份／暫存／舊檔）。ZW3D 安裝目錄實測：
# 備份檔本身 *.z3bak 與語言無關，差異在使用者自建的目錄／檔名用語。
# 這些詞只在「名稱或路徑已與 ZW3D 相關」時才作為附加命中條件，且排除安裝目錄。
$script:LangWords = @{
 'cs_CZ' = @('záloha', 'dočasný', 'starý')
 'de_DE' = @('sicherung', 'sicherungskopie', 'temporär', 'alt')
 'en_US' = @('backup', 'autosave', 'temp', 'old', 'cache')
 'es_ES' = @('copia', 'respaldo', 'temporal', 'antiguo')
 'fr_FR' = @('sauvegarde', 'temporaire', 'ancien')
 'hu_HU' = @('mentés', 'biztonsági', 'ideiglenes', 'régi')
 'it_IT' = @('backup', 'copia', 'temporaneo', 'vecchio')
 'ja_JP' = @('バックアップ', '一時', '古い', '自動保存')
 'ko_KR' = @('백업', '임시', '오래된', '자동 저장')
 'pl_PL' = @('kopia', 'zapasowa', 'tymczasowy', 'stary')
 'pt_PT' = @('cópia', 'cópia de segurança', 'temporário', 'antigo')
 'ru_RU' = @('резервная', 'копия', 'временный', 'старый')
 'tr_TR' = @('yedek', 'yedekleme', 'geçici', 'eski')
 'zh_CN' = @('备份', '备份文件', '暂存', '临时', '旧', '自动保存')
 'zh_TW' = @('備份', '備份檔', '暫存', '舊', '自動儲存')
}

# 空 Id = 全部語言（所有關鍵詞）；否則只取該語言組。
function Get-LangWords([string]$Id) {
    if ([string]::IsNullOrEmpty($Id)) {
        $all = New-Object System.Collections.ArrayList
        foreach ($l in $script:Langs) {
            foreach ($w in $script:LangWords[$l.Id]) {
                $lw = $w.ToLowerInvariant()
                if (-not $all.Contains($lw)) { [void]$all.Add($lw) }
            }
        }
        return $all.ToArray()
    }
    return @($script:LangWords[$Id] | ForEach-Object { $_.ToLowerInvariant() })
}

# 常用關鍵字預設：保留原有英文詞，再依目前語言版本補上在地化詞
function Get-PresetWords([string]$Id) {
    $base = @('ZW3D', 'Temp', 'Backup', 'Cache', 'log', 'old')
    $extra = New-Object System.Collections.ArrayList
    if (-not [string]::IsNullOrEmpty($Id)) {
        foreach ($w in $script:LangWords[$Id]) { [void]$extra.Add($w) }
    }
    $out = New-Object System.Collections.ArrayList
    foreach ($w in (@($base) + $extra.ToArray())) {
        if (-not $out.Contains($w)) { [void]$out.Add($w) }
    }
    return $out.ToArray()
}

# 偏好：exe 同目錄 language.txt（兩行：ui=介面語言 / scan=掃描語言版本）。
# 相容 v1.4.0 舊格式（單一語言碼）：視為兩個設定套用同一語言。
function Save-LangPref {
    $txt = 'ui={0}{1}scan={2}' -f $script:UiLang, [Environment]::NewLine, $script:ScanLangId
    try { [System.IO.File]::WriteAllText($script:LangFile, $txt, (New-Object System.Text.UTF8Encoding($false))) } catch { }
}
function Test-LangId([string]$Id) {
    if ([string]::IsNullOrEmpty($Id)) { return $true }
    return (@($script:Langs | Where-Object { $_.Id -eq $Id }).Count -gt 0)
}
try {
    if (Test-Path -LiteralPath $script:LangFile) {
        $lines = @([System.IO.File]::ReadAllLines($script:LangFile))
        $isNew = $false
        foreach ($ln in $lines) {
            $t = ([string]$ln).Trim()
            if ($t.StartsWith('ui=')) {
                $v = $t.Substring(3).Trim()
                if ($v -and (Test-LangId $v)) { $script:UiLang = $v }
                $isNew = $true
            } elseif ($t.StartsWith('scan=')) {
                $v = $t.Substring(5).Trim()
                if (Test-LangId $v) { $script:ScanLangId = $v }
                $isNew = $true
            }
        }
        if (-not $isNew -and $lines.Count -gt 0) {
            $old = ([string]$lines[0]).Trim()
            if (Test-LangId $old) {
                $script:ScanLangId = $old
                if ($old) { $script:UiLang = $old } else { $script:UiLang = 'zh_TW' }
            }
        }
    }
} catch { }

function Get-UiFontFamily {
    $l = $script:Langs | Where-Object { $_.Id -eq $script:UiLang } | Select-Object -First 1
    if ($l -and $l.Font) { return $l.Font }
    return 'Microsoft JhengHei UI'
}

function Flush-Log {
    if ($script:LogFilePending.Length -eq 0 -and $script:LogUiPending.Length -eq 0) { return }
    if ($script:LogFilePending.Length -gt 0) {
        try { [System.IO.File]::AppendAllText($script:LogFile, $script:LogFilePending.ToString(), $script:Utf8) } catch { }
        [void]$script:LogFilePending.Clear()
    }
    if ($script:LogUiPending.Length -gt 0) {
        try {
            if ($script:LogBox -and -not $script:LogBox.IsDisposed) {
                $script:LogBox.AppendText($script:LogUiPending.ToString())
                # 用計數器判斷，不再每次呼叫 LogBox.Lines（會重建整份字串陣列，O(n)）
                if ($script:LogLineCount -gt 3000) {
                    $arr = $script:LogBox.Lines
                    if ($arr.Count -gt 1500) {
                        $script:LogBox.Text = ($arr[-1500..-1] -join "`r`n") + "`r`n"
                        $script:LogLineCount = 1500
                    }
                }
            }
        } catch { }
        [void]$script:LogUiPending.Clear()
    }
    $script:LogClock.Restart()
}

function Write-Log {
    param([string]$Level, [string]$Message)
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    [void]$script:LogFilePending.Append($line + [Environment]::NewLine)
    [void]$script:LogUiPending.Append($line + "`r`n")
    $script:LogLineCount++
    if ($script:LogBatching) {
        if (($script:LogLineCount % 100) -eq 0 -or $script:LogClock.ElapsedMilliseconds -ge 250) { Flush-Log }
    } elseif (($script:LogLineCount % 10) -eq 0 -or $script:LogClock.ElapsedMilliseconds -ge 100) {
        Flush-Log
    }
}

# 一般模式最多延遲 150ms 顯示；關閉與大量刪除時另有明確 Flush-Log
$script:LogTimer = New-Object System.Windows.Forms.Timer
$script:LogTimer.Interval = 150
$script:LogTimer.Add_Tick({ Flush-Log })
$script:LogTimer.Start()

# ---------------------------------------------------------------- 色彩 / 字型
$C_BG      = [System.Drawing.Color]::FromArgb(28, 28, 30)
$C_PANEL   = [System.Drawing.Color]::FromArgb(38, 38, 42)
$C_FIELD   = [System.Drawing.Color]::FromArgb(48, 48, 54)
$C_TEXT    = [System.Drawing.Color]::FromArgb(232, 232, 235)
$C_MUTED   = [System.Drawing.Color]::FromArgb(150, 152, 160)
$C_ACCENT  = [System.Drawing.Color]::FromArgb(0, 120, 212)
$C_DANGER  = [System.Drawing.Color]::FromArgb(196, 62, 62)
$C_WARN    = [System.Drawing.Color]::FromArgb(232, 168, 64)
$C_OK      = [System.Drawing.Color]::FromArgb(78, 176, 110)

$FONT_UI   = New-Object System.Drawing.Font('Microsoft JhengHei UI', 10)
$FONT_BOLD = New-Object System.Drawing.Font('Microsoft JhengHei UI', 10, [System.Drawing.FontStyle]::Bold)
$FONT_MONO = New-Object System.Drawing.Font('Consolas', 9)

# ---------------------------------------------------------------- 副檔名預設清單
# 要擴充選項：直接在這個陣列加字串即可（例：'.step'）；下拉選單也可直接手動輸入任意副檔名。
$script:ExtPresets = @(
    '（不限）',
    '.z3', '.z3prt', '.z3asm', '.z3drw', '.z3bak', '.z3gr',
    '.xml', '.log', '.cfg', '.info', '.orig',
    '.txt', '.csv', '.json', '.stl', '.bmp', '.lim', '.ptss', '.job', '.tls', '.wkp',
    '.projector', '.offsetrep', '.vxbcl'
)
# 匯出清單支援的格式
$script:ExportExts = @('.txt', '.csv', '.json')

function New-DarkButton {
    param([string]$Text, [int]$Width = 110, [int]$Height = 34, $BackColor = $C_FIELD, $ForeColor = $C_TEXT)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Width = $Width
    $b.Height = $Height
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(70, 70, 78)
    $b.FlatAppearance.BorderSize = 1
    $b.BackColor = $BackColor
    $b.ForeColor = $ForeColor
    $b.Font = $FONT_BOLD
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    $b.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    return $b
}

# ---------------------------------------------------------------- 大圖示元件
$script:Tip = New-Object System.Windows.Forms.ToolTip
$script:Tip.AutoPopDelay = 9000
$script:Tip.InitialDelay = 200
$script:Tip.ReshowDelay = 80
$script:Tip.ShowAlways = $true

# ---------------------------------------------------------------- 向量圖示（32x32 格，亮藍實心風格）
# 填色 even-odd（重疊處挖空）：c cx cy r | r x y w h rad | p x y x y ...；描邊圓頭：l x1 y1 x2 y2 w | a cx cy r start sweep w
$C_ICON = [System.Drawing.Color]::FromArgb(92, 176, 255)
$C_ICON2 = [System.Drawing.Color]::FromArgb(206, 232, 255)
$C_MUTED2 = [System.Drawing.Color]::FromArgb(96, 98, 106)
$script:IconData = @{
    search = @('c 13.5 13.5 10','c 13.5 13.5 6.3','l 21.2 21.2 27.5 27.5 5','|','c 13.5 13.5 5.7')
    stop = @('c 16 16 13.5','|','p 19 10.6 21.4 13 18.4 16 21.4 19 19 21.4 16 18.4 13 21.4 10.6 19 13.6 16 10.6 13 13 10.6 16 13.6')
    checkall = @('r 3 3 26 26 8','|','p 8.6 16.4 10.9 14.1 14 17.2 21.2 10 23.4 12.3 14 21.7')
    uncheck = @('r 3 3 26 26 8','r 6.8 6.8 18.4 18.4 4.5','|','p 18.4 11.6 20.4 13.6 18 16 20.4 18.4 18.4 20.4 16 18 13.6 20.4 11.6 18.4 14 16 11.6 13.6 13.6 11.6 16 14')
    delete = @('r 4 6 24 4 2','r 12 2.6 8 3.4 1.7','p 7 12.6 25 12.6 23.6 27.2 23.1 28.4 22 29 10 29 8.9 28.4 8.4 27.2','|','r 12.2 15.5 2.4 9.5 1.2','r 17.4 15.5 2.4 9.5 1.2')
    folder = @('r 3 6 26 22 4.5','r 3 3.3 12 3 1.4','|','r 3 11.5 26 16.5 4.5')
    log = @('r 5.5 3 21 26 5','|','r 10 10.5 12 2.4 1.2','r 10 15.4 12 2.4 1.2','r 10 20.3 7.5 2.4 1.2')
    export = @('p 4 19 7.5 19 7.5 23.5 24.5 23.5 24.5 19 28 19 28 25 26.5 28.5 5.5 28.5 4 25','|','p 16 3 23 10 18 10 18 20 14 20 14 10 9 10')
    list = @('r 3.5 4.5 6.4 6.4 2','r 3.5 12.8 6.4 6.4 2','r 3.5 21.1 6.4 6.4 2','|','l 14 7.7 28 7.7 3','l 14 16 28 16 3','l 14 24.3 28 24.3 3')
    power = @('a 16 17 10.2 -50 280 4','|','l 16 3.8 16 15 4')
    refresh = @('a 16 16 10.4 -20 290 4','|','p 15.6 0.6 24.8 5.6 15.6 10.6')
}
function New-Glyph {
    param([string]$Name, [int]$Px, $C1, $C2)
    $u = $Px / 32.0
    $bmp = New-Object System.Drawing.Bitmap($Px, $Px)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $P = @((New-Object System.Drawing.Drawing2D.GraphicsPath), (New-Object System.Drawing.Drawing2D.GraphicsPath))
    $C = @($C1, $C2)
    $k = 0
    foreach ($o in $script:IconData[$Name]) {
        if ($o -eq '|') { $k = 1; continue }
        $path = $P[$k]
        $t = $o.Split(' ')
        $v = @($t[1..($t.Count - 1)] | ForEach-Object { [single]$_ })
        switch ($t[0]) {
            'c' { $path.AddEllipse(($v[0] - $v[2]) * $u, ($v[1] - $v[2]) * $u, 2 * $v[2] * $u, 2 * $v[2] * $u) }
            'r' {
                $d = 2 * $v[4] * $u
                $x = $v[0] * $u; $y = $v[1] * $u; $w = $v[2] * $u; $h = $v[3] * $u
                $path.StartFigure()
                $path.AddArc($x, $y, $d, $d, 180, 90)
                $path.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
                $path.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
                $path.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
                $path.CloseFigure()
            }
            'p' {
                $pts = for ($i = 0; $i -lt $v.Count; $i += 2) { New-Object System.Drawing.PointF(($v[$i] * $u), ($v[$i + 1] * $u)) }
                $path.AddPolygon([System.Drawing.PointF[]]@($pts))
            }
            default {
                $pen = New-Object System.Drawing.Pen $C[$k], ($v[$v.Count - 1] * $u)
                $pen.StartCap = 'Round'; $pen.EndCap = 'Round'
                if ($t[0] -eq 'l') { $g.DrawLine($pen, $v[0] * $u, $v[1] * $u, $v[2] * $u, $v[3] * $u) }
                else { $g.DrawArc($pen, ($v[0] - $v[2]) * $u, ($v[1] - $v[2]) * $u, 2 * $v[2] * $u, 2 * $v[2] * $u, $v[3], $v[4]) }
                $pen.Dispose()
            }
        }
    }
    foreach ($k in 0, 1) {
        $br = New-Object System.Drawing.SolidBrush $C[$k]
        $pen = New-Object System.Drawing.Pen $C[$k], ($u * 1.1)
        $pen.LineJoin = 'Round'
        $g.FillPath($br, $P[$k])
        $g.DrawPath($pen, $P[$k])
        $br.Dispose(); $pen.Dispose(); $P[$k].Dispose()
    }
    $g.Dispose()
    return $bmp
}

function Get-HoverColor {
    param($c)
    return [System.Drawing.Color]::FromArgb(
        [math]::Min(255, $c.R + 26),
        [math]::Min(255, $c.G + 26),
        [math]::Min(255, $c.B + 26))
}

# 大圖示按鈕：上圖示 + 下文字 + hover 變色 + Tooltip
function New-IconButton {
    param([string]$Icon, [string]$CapKey, [string]$TipKey,
          [int]$Width = 96, [int]$Height = 94,
          $BackColor = $C_FIELD, $ForeColor = $C_TEXT)

    $b = New-Object System.Windows.Forms.Button
    $b.Width = $Width
    $b.Height = $Height
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(70, 70, 78)
    $b.FlatAppearance.BorderSize = 1
    $b.BackColor = $BackColor
    $b.ForeColor = $ForeColor
    $b.Cursor = [System.Windows.Forms.Cursors]::Hand
    $b.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
    $b.Tag = $BackColor

    $px = 36; if ($Height -lt 80) { $px = 30 }
    $g = New-Object System.Windows.Forms.Label
    $g.Dock = 'Fill'
    $g.ImageAlign = 'MiddleCenter'
    $g.BackColor = [System.Drawing.Color]::Transparent
    if ($ForeColor.R -eq 255 -and $ForeColor.B -eq 255) {
        $imE = New-Glyph $Icon $px $ForeColor ([System.Drawing.Color]::FromArgb([int](($BackColor.R + 255) / 2), [int](($BackColor.G + 255) / 2), [int](($BackColor.B + 255) / 2)))
    } else { $imE = New-Glyph $Icon $px $C_ICON $C_ICON2 }
    $imD = New-Glyph $Icon $px $C_MUTED $C_MUTED2
    $g.Image = $imE

    $c = New-Object System.Windows.Forms.Label
    $c.Text = L $CapKey
    $c.Font = New-Object System.Drawing.Font('Microsoft JhengHei UI', 8.5)
    $c.ForeColor = $C_MUTED
    $c.Dock = 'Bottom'
    $c.Height = 20
    $c.TextAlign = 'MiddleCenter'
    $c.BackColor = [System.Drawing.Color]::Transparent

    $b.Controls.Add($g)
    $b.Controls.Add($c)   # 後加的先 Dock：caption 先佔底部，圖示佔滿其餘

    foreach ($ctrl in @($b, $g, $c)) {
        $ctrl.Tag = @{ B = $b; C = $BackColor; I = $g; E = $imE; D = $imD }
        $ctrl.Add_MouseEnter({ $x = $args[0].Tag; $x.B.BackColor = Get-HoverColor $x.C })
        $ctrl.Add_MouseLeave({ $x = $args[0].Tag; $x.B.BackColor = $x.C })
        $script:Tip.SetToolTip($ctrl, (L $TipKey))
    }
    # 只有圖示/文字標籤需要把點擊轉發給按鈕；按鈕自身再掛轉發會無限遞迴
    $g.Add_Click({ $args[0].Tag.B.PerformClick() })
    $c.Add_Click({ $args[0].Tag.B.PerformClick() })
    $b.Add_EnabledChanged({ $x = $args[0].Tag; if ($x.B.Enabled) { $x.I.Image = $x.E } else { $x.I.Image = $x.D } })
    $b | Add-Member -NotePropertyName CapLabel -NotePropertyValue $c -Force
    $b | Add-Member -NotePropertyName CapKey   -NotePropertyValue $CapKey -Force
    $b | Add-Member -NotePropertyName TipKey   -NotePropertyValue $TipKey -Force
    [void]$script:IconButtons.Add($b)
    return $b
}

# 圖示列的群組分隔線
function New-RibbonSeparator {
    param($parent)
    $p = New-Object System.Windows.Forms.Panel
    $p.Width = 1
    $p.Height = 80
    $p.Margin = New-Object System.Windows.Forms.Padding(4, 7, 4, 0)
    $p.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 66)
    $parent.Controls.Add($p)
}

# ---------------------------------------------------------------- 主視窗
$form = New-Object System.Windows.Forms.Form
$form.Text = L 'AppTitle'
$form.Size = New-Object System.Drawing.Size(1340, 880)
$form.MinimumSize = New-Object System.Drawing.Size(1200, 760)
$form.StartPosition = 'CenterScreen'
$form.BackColor = $C_BG
$form.ForeColor = $C_TEXT
$form.Font = $FONT_UI
$form.KeyPreview = $true

# App 圖示（取自 exe 內嵌 app.ico）：標題列、工作列、Alt+Tab 都吃這顆
try {
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
} catch { Write-Log 'WARN' ((L 'Log.IconFail') -f $_.Exception.Message) }

# ---------- 結果清單（先加，Dock Fill）
$lvResults = New-Object System.Windows.Forms.ListView
$lvResults.Dock = 'Fill'
$lvResults.View = 'Details'
$lvResults.FullRowSelect = $true
$lvResults.CheckBoxes = $true
$lvResults.GridLines = $false
$lvResults.HideSelection = $false
$lvResults.MultiSelect = $true
$lvResults.BorderStyle = 'None'
$lvResults.BackColor = [System.Drawing.Color]::FromArgb(24, 24, 26)
$lvResults.ForeColor = $C_TEXT
$lvResults.Font = $FONT_UI
$lvResults.ShowItemToolTips = $false   # 5000 列時 hover 命中測試會造成拖動卡頓；完整路徑已在第三欄
try {
    $pi = $lvResults.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'Instance,NonPublic')
    if ($pi) { $pi.SetValue($lvResults, $true, $null) }
} catch { }
[void]$lvResults.Columns.Add((L 'Col.Name'), 240)
[void]$lvResults.Columns.Add((L 'Col.Type'), 70)
[void]$lvResults.Columns.Add((L 'Col.Path'), 460)
[void]$lvResults.Columns.Add((L 'Col.Parent'), 230)
$form.Controls.Add($lvResults)

# ---------- 頂部工具列
$pnlTop = New-Object System.Windows.Forms.Panel
$pnlTop.Dock = 'Top'
$pnlTop.Height = 150
$pnlTop.BackColor = $C_PANEL
$pnlTop.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 8)
$form.Controls.Add($pnlTop)

$lblKw = New-Object System.Windows.Forms.Label
$lblKw.Text = L 'KwLabel'
$lblKw.AutoSize = $true
$lblKw.Location = New-Object System.Drawing.Point(14, 12)
$lblKw.ForeColor = $C_MUTED
$pnlTop.Controls.Add($lblKw)

$txtKeyword = New-Object System.Windows.Forms.TextBox
$txtKeyword.Location = New-Object System.Drawing.Point(14, 34)
$txtKeyword.Width = 420
$txtKeyword.Height = 30
$txtKeyword.BackColor = $C_FIELD
$txtKeyword.ForeColor = $C_TEXT
$txtKeyword.BorderStyle = 'FixedSingle'
$txtKeyword.Font = New-Object System.Drawing.Font('Microsoft JhengHei UI', 11)
$pnlTop.Controls.Add($txtKeyword)

$lblPreset = New-Object System.Windows.Forms.Label
$lblPreset.Text = L 'PresetLabel'
$lblPreset.AutoSize = $true
$lblPreset.Location = New-Object System.Drawing.Point(448, 12)
$lblPreset.ForeColor = $C_MUTED
$pnlTop.Controls.Add($lblPreset)

$cboPreset = New-Object System.Windows.Forms.ComboBox
$cboPreset.DropDownStyle = 'DropDown'
$cboPreset.Location = New-Object System.Drawing.Point(448, 34)
$cboPreset.Width = 150
$cboPreset.BackColor = $C_FIELD
$cboPreset.ForeColor = $C_TEXT
$cboPreset.FlatStyle = 'Flat'
[void]$cboPreset.Items.AddRange(@((L 'Preset.All'), 'ZW3D', 'Temp', 'Backup', 'Cache', 'log', 'old'))
$cboPreset.SelectedIndex = 0
$pnlTop.Controls.Add($cboPreset)

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = L 'TargetLabel'
$lblTarget.AutoSize = $true
$lblTarget.Location = New-Object System.Drawing.Point(614, 12)
$lblTarget.ForeColor = $C_MUTED
$pnlTop.Controls.Add($lblTarget)

$cboTarget = New-Object System.Windows.Forms.ComboBox
$cboTarget.DropDownStyle = 'DropDownList'
$cboTarget.Location = New-Object System.Drawing.Point(614, 34)
$cboTarget.Width = 110
$cboTarget.BackColor = $C_FIELD
$cboTarget.ForeColor = $C_TEXT
$cboTarget.FlatStyle = 'Flat'
[void]$cboTarget.Items.AddRange(@((L 'Target.Folder'), (L 'Target.File'), (L 'Target.Both')))
$cboTarget.SelectedIndex = 1
$pnlTop.Controls.Add($cboTarget)

$lblDepth = New-Object System.Windows.Forms.Label
$lblDepth.Text = L 'DepthLabel'
$lblDepth.AutoSize = $true
$lblDepth.Location = New-Object System.Drawing.Point(740, 12)
$lblDepth.ForeColor = $C_MUTED
$pnlTop.Controls.Add($lblDepth)

$cboDepth = New-Object System.Windows.Forms.ComboBox
$cboDepth.DropDownStyle = 'DropDownList'
$cboDepth.Location = New-Object System.Drawing.Point(740, 34)
$cboDepth.Width = 90
$cboDepth.BackColor = $C_FIELD
$cboDepth.ForeColor = $C_TEXT
$cboDepth.FlatStyle = 'Flat'
[void]$cboDepth.Items.AddRange(@((L 'Depth.3'), (L 'Depth.5'), (L 'Depth.8'), (L 'Depth.12'), (L 'Depth.Unlimited')))
$cboDepth.SelectedIndex = 3
$pnlTop.Controls.Add($cboDepth)

$chkPermanent = New-Object System.Windows.Forms.CheckBox
$chkPermanent.Text = L 'PermanentChk'
$chkPermanent.AutoSize = $true
$chkPermanent.Location = New-Object System.Drawing.Point(846, 38)
$chkPermanent.ForeColor = $C_WARN
$pnlTop.Controls.Add($chkPermanent)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = L 'Hint'
$lblHint.AutoSize = $true
$lblHint.Location = New-Object System.Drawing.Point(14, 122)
$lblHint.ForeColor = $C_MUTED
$lblHint.Font = New-Object System.Drawing.Font('Microsoft JhengHei UI', 9)
$pnlTop.Controls.Add($lblHint)

# ---------- ZW3D 篩選（第二列）
$chkZw3dOnly = New-Object System.Windows.Forms.CheckBox
$chkZw3dOnly.Text = L 'Zw3dOnlyChk'
$chkZw3dOnly.AutoSize = $true
$chkZw3dOnly.Location = New-Object System.Drawing.Point(14, 82)
$chkZw3dOnly.ForeColor = $C_TEXT
$pnlTop.Controls.Add($chkZw3dOnly)

$cboZw3dScope = New-Object System.Windows.Forms.ComboBox
$cboZw3dScope.DropDownStyle = 'DropDownList'
$cboZw3dScope.Location = New-Object System.Drawing.Point(190, 80)
$cboZw3dScope.Width = 250
$cboZw3dScope.BackColor = $C_FIELD
$cboZw3dScope.ForeColor = $C_TEXT
$cboZw3dScope.FlatStyle = 'Flat'
$cboZw3dScope.Enabled = $false
[void]$cboZw3dScope.Items.AddRange(@((L 'ZwScope.All'), (L 'ZwScope.Junk'), (L 'ZwScope.Model')))
$cboZw3dScope.SelectedIndex = 0
$pnlTop.Controls.Add($cboZw3dScope)

$lblExt = New-Object System.Windows.Forms.Label
$lblExt.Text = L 'ExtLabel'
$lblExt.AutoSize = $true
$lblExt.Location = New-Object System.Drawing.Point(446, 84)
$lblExt.ForeColor = $C_MUTED
$pnlTop.Controls.Add($lblExt)

$cboExt = New-Object System.Windows.Forms.ComboBox
$cboExt.DropDownStyle = 'DropDown'      # 可編輯 → 清單可擴充、也能臨時輸入
$cboExt.Location = New-Object System.Drawing.Point(506, 80)
$cboExt.Width = 130
$cboExt.BackColor = $C_FIELD
$cboExt.ForeColor = $C_TEXT
$cboExt.FlatStyle = 'Flat'
[void]$cboExt.Items.AddRange($script:ExtPresets)
$cboExt.SelectedIndex = 0
$pnlTop.Controls.Add($cboExt)

# 搜尋對象＝資料夾時，副檔名不適用 → 停用並顯示（不限）
function Update-ExtEnabled {
    $isFolder = ($cboTarget.SelectedIndex -eq 0)
    $cboExt.Enabled = -not $isFolder
    if ($isFolder -and $cboExt.Items.Count -gt 0) { $cboExt.SelectedIndex = 0 }
}
$cboTarget.Add_SelectedIndexChanged({ Update-ExtEnabled })

$lblZw3dHint = New-Object System.Windows.Forms.Label
$lblZw3dHint.Text = L 'Zw3dHint'
$lblZw3dHint.AutoSize = $true
$lblZw3dHint.Location = New-Object System.Drawing.Point(648, 84)
$lblZw3dHint.ForeColor = $C_MUTED
$lblZw3dHint.Font = New-Object System.Drawing.Font('Microsoft JhengHei UI', 9)

# ---------- 介面語言 / ZW3D 語言版本（頂部第三列右側；兩個設定互相獨立）
# 介面語言：只切換本工具文案與字體，不影響 ZW3D 軟體，也不影響搜尋結果
$lblUi = New-Object System.Windows.Forms.Label
$lblUi.Text = L 'UiLangLabel'
$lblUi.AutoSize = $false
$lblUi.Size = New-Object System.Drawing.Size(140, 20)
$lblUi.Location = New-Object System.Drawing.Point(560, 116)
$lblUi.ForeColor = $C_MUTED
$pnlTop.Controls.Add($lblUi)

$cboUi = New-Object System.Windows.Forms.ComboBox
$cboUi.DropDownStyle = 'DropDownList'
$cboUi.Location = New-Object System.Drawing.Point(702, 112)
$cboUi.Width = 150
$cboUi.BackColor = $C_FIELD
$cboUi.ForeColor = $C_TEXT
$cboUi.FlatStyle = 'Flat'
$cboUi.Font = $FONT_UI
foreach ($l in $script:Langs) { [void]$cboUi.Items.Add(('{0}  {1}' -f $l.Id, $l.Name)) }
$cboUi.SelectedIndex = 14                      # 預設繁體中文
for ($i = 0; $i -lt $script:Langs.Count; $i++) {
    if ($script:Langs[$i].Id -eq $script:UiLang) { $cboUi.SelectedIndex = $i; break }
}
$pnlTop.Controls.Add($cboUi)
$script:Tip.SetToolTip($cboUi, (L 'Tip.UiLang'))
$script:Tip.SetToolTip($lblUi, (L 'Tip.UiLang'))

# ZW3D 語言版本：只影響常用關鍵字與掃描時的在地化關鍵詞，不影響介面文案
$lblLang = New-Object System.Windows.Forms.Label
$lblLang.Text = L 'LangLabel'
$lblLang.AutoSize = $false
$lblLang.Size = New-Object System.Drawing.Size(160, 20)
$lblLang.Location = New-Object System.Drawing.Point(868, 116)
$lblLang.ForeColor = $C_MUTED
$pnlTop.Controls.Add($lblLang)

$cboLang = New-Object System.Windows.Forms.ComboBox
$cboLang.DropDownStyle = 'DropDownList'
$cboLang.Location = New-Object System.Drawing.Point(1030, 112)
$cboLang.Width = 205
$cboLang.BackColor = $C_FIELD
$cboLang.ForeColor = $C_TEXT
$cboLang.FlatStyle = 'Flat'
$cboLang.Font = $FONT_UI
[void]$cboLang.Items.Add((L 'Lang.All'))
foreach ($l in $script:Langs) { [void]$cboLang.Items.Add(('{0}  {1}' -f $l.Id, $l.Name)) }
$cboLang.SelectedIndex = 0
if ($script:ScanLangId) {
    for ($i = 0; $i -lt $script:Langs.Count; $i++) {
        if ($script:Langs[$i].Id -eq $script:ScanLangId) { $cboLang.SelectedIndex = $i + 1; break }
    }
}
$pnlTop.Controls.Add($cboLang)
$script:Tip.SetToolTip($cboLang, (L 'Tip.ScanLang'))
$script:Tip.SetToolTip($lblLang, (L 'Tip.ScanLang'))
$pnlTop.Controls.Add($lblZw3dHint)

$chkZw3dOnly.Add_CheckedChanged({ $cboZw3dScope.Enabled = $chkZw3dOnly.Checked })
# 預設開啟：本工具的用途就是找 ZW3D 相關檔案；取消勾選即可回到全檔案搜尋
$chkZw3dOnly.Checked = $true
$cboZw3dScope.Enabled = $true

# ---------- 大圖示工具列（Ribbon，會隨視窗寬度自動換行）
$pnlRibbon = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlRibbon.Dock = 'Top'
$pnlRibbon.AutoSize = $true
$pnlRibbon.AutoSizeMode = 'GrowAndShrink'
$pnlRibbon.MinimumSize = New-Object System.Drawing.Size(0, 112)
$pnlRibbon.BackColor = [System.Drawing.Color]::FromArgb(32, 32, 35)
$pnlRibbon.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 8)
$pnlRibbon.WrapContents = $true
$form.Controls.Add($pnlRibbon)

$btnSearch = New-IconButton -Icon 'search' -CapKey 'Btn.Search' -BackColor $C_ACCENT -ForeColor ([System.Drawing.Color]::White) `
    -TipKey 'Tip.Search'
$pnlRibbon.Controls.Add($btnSearch)
$btnStop = New-IconButton -Icon 'stop' -CapKey 'Btn.Stop' `
    -TipKey 'Tip.Stop'
$btnStop.Enabled = $false
$pnlRibbon.Controls.Add($btnStop)
New-RibbonSeparator $pnlRibbon

$btnCheckResults = New-IconButton -Icon 'checkall' -CapKey 'Btn.CheckAll' `
    -TipKey 'Tip.CheckAll'
$pnlRibbon.Controls.Add($btnCheckResults)
$btnUncheckResults = New-IconButton -Icon 'uncheck' -CapKey 'Btn.Uncheck' `
    -TipKey 'Tip.Uncheck'
$pnlRibbon.Controls.Add($btnUncheckResults)
New-RibbonSeparator $pnlRibbon

$btnDelete = New-IconButton -Icon 'delete' -CapKey 'Btn.Delete' -BackColor $C_DANGER -ForeColor ([System.Drawing.Color]::White) `
    -TipKey 'Tip.Delete'
$pnlRibbon.Controls.Add($btnDelete)
New-RibbonSeparator $pnlRibbon

$btnOpenFolder = New-IconButton -Icon 'folder' -CapKey 'Btn.Explorer' `
    -TipKey 'Tip.Explorer'
$pnlRibbon.Controls.Add($btnOpenFolder)
$btnOpenLog = New-IconButton -Icon 'log' -CapKey 'Btn.LogFolder' `
    -TipKey 'Tip.LogFolder'
$pnlRibbon.Controls.Add($btnOpenLog)
$btnExport = New-IconButton -Icon 'export' -CapKey 'Btn.Export' `
    -TipKey 'Tip.Export'
$pnlRibbon.Controls.Add($btnExport)
$btnUserFolder = New-IconButton -Icon 'folder' -CapKey 'Btn.UserFolder' `
    -TipKey 'Tip.UserFolder'
$pnlRibbon.Controls.Add($btnUserFolder)
New-RibbonSeparator $pnlRibbon

$btnAbort = New-IconButton -Icon 'power' -CapKey 'Btn.Abort' -BackColor $C_DANGER -ForeColor ([System.Drawing.Color]::White) `
    -TipKey 'Tip.Abort'
$pnlRibbon.Controls.Add($btnAbort)

# ---------- 左側磁碟機
$pnlLeft = New-Object System.Windows.Forms.Panel
$pnlLeft.Dock = 'Left'
$pnlLeft.Width = 320
$pnlLeft.BackColor = $C_PANEL
$pnlLeft.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)
$form.Controls.Add($pnlLeft)

$lblDrives = New-Object System.Windows.Forms.Label
$lblDrives.Text = L 'DrivesLabel'
$lblDrives.Dock = 'Top'
$lblDrives.Height = 26
$lblDrives.ForeColor = $C_MUTED
$pnlLeft.Controls.Add($lblDrives)

$clbDrives = New-Object System.Windows.Forms.CheckedListBox
$clbDrives.Dock = 'Fill'
$clbDrives.CheckOnClick = $true
$clbDrives.BorderStyle = 'None'
$clbDrives.BackColor = [System.Drawing.Color]::FromArgb(24, 24, 26)
$clbDrives.ForeColor = $C_TEXT
$clbDrives.Font = New-Object System.Drawing.Font('Microsoft JhengHei UI', 9)
$clbDrives.IntegralHeight = $false
$clbDrives.DisplayMember = 'Text'
$pnlLeft.Controls.Add($clbDrives)
$clbDrives.BringToFront()

$pnlDriveBtns = New-Object System.Windows.Forms.FlowLayoutPanel
$pnlDriveBtns.Dock = 'Bottom'
$pnlDriveBtns.Height = 76
$pnlDriveBtns.BackColor = $C_PANEL
$pnlDriveBtns.WrapContents = $true
$pnlDriveBtns.Padding = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
$pnlLeft.Controls.Add($pnlDriveBtns)

$btnRefresh = New-IconButton -Icon 'refresh' -CapKey 'Btn.Refresh' -Width 132 -Height 64 `
    -TipKey 'Tip.Refresh'
$pnlDriveBtns.Controls.Add($btnRefresh)
$btnCheckAll = New-IconButton -Icon 'list' -CapKey 'Btn.ToggleDrives' -Width 132 -Height 64 `
    -TipKey 'Tip.ToggleDrives'
$pnlDriveBtns.Controls.Add($btnCheckAll)

# ---------- 底部：按鈕 + 日誌
$pnlBottom = New-Object System.Windows.Forms.Panel
$pnlBottom.Dock = 'Bottom'
$pnlBottom.Height = 200
$pnlBottom.BackColor = $C_PANEL
$pnlBottom.Padding = New-Object System.Windows.Forms.Padding(10, 8, 10, 8)
$form.Controls.Add($pnlBottom)

$script:LogBox = New-Object System.Windows.Forms.TextBox
$script:LogBox.Multiline = $true
$script:LogBox.Dock = 'Fill'
$script:LogBox.ReadOnly = $true
$script:LogBox.ScrollBars = 'Vertical'
$script:LogBox.BackColor = [System.Drawing.Color]::FromArgb(18, 18, 20)
$script:LogBox.ForeColor = [System.Drawing.Color]::FromArgb(170, 200, 180)
$script:LogBox.Font = $FONT_MONO
$script:LogBox.BorderStyle = 'None'
$script:LogBox.WordWrap = $false
$pnlBottom.Controls.Add($script:LogBox)
$script:LogBox.BringToFront()

# ---------- 狀態列
$status = New-Object System.Windows.Forms.StatusStrip
$status.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 22)
$status.ForeColor = $C_MUTED
$status.SizingGrip = $false
$lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblStatus.Text = L 'Status.Ready'
$lblStatus.Spring = $true
$lblStatus.TextAlign = 'MiddleLeft'
$lblScan = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblScan.Text = ''
$lblScan.ForeColor = $C_MUTED
[void]$status.Items.Add($lblStatus)
[void]$status.Items.Add($lblScan)
# 開發者標註：狀態列最右端（lblStatus Spring 會把它推到右邊）
$lblCredit = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblCredit.Text = 'Developed by 阿家_diabloz'
$lblCredit.ForeColor = $C_MUTED
$lblCredit.ToolTipText = L 'CreditTip'
[void]$status.Items.Add($lblCredit)
$form.Controls.Add($status)

# ---------- 刪除進度條（狀態列上方，平時隱藏）
# WinForms 內建 ProgressBar 無法自訂顏色，改自繪：底色軌道 + 彩色填充 + 百分比文字
$pnlProg = New-Object System.Windows.Forms.Panel
$pnlProg.Dock = 'Bottom'
$pnlProg.Height = 0
$pnlProg.BackColor = [System.Drawing.Color]::FromArgb(32, 32, 36)
$script:ProgPct = 0
$pnlProg.Add_Paint({
    $g = $args[1].Graphics
    $g.SmoothingMode = 'AntiAlias'
    $w = $args[0].Width - 24
    $h = 10
    $x = 12
    $y = 8
    # 軌道
    $track = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(52, 52, 58))
    $g.FillRectangle($track, $x, $y, $w, $h)
    # 填充（依百分比）
    $fw = [int]($w * $script:ProgPct / 100)
    if ($fw -gt 0) {
        $fill = New-Object System.Drawing.SolidBrush ($script:ProgColor)
        $g.FillRectangle($fill, $x, $y, $fw, $h)
        $fill.Dispose()
    }
    $track.Dispose()
    # 百分比文字（置中於整條軌道，固定可見）
    $txt = '{0}%' -f $script:ProgPct
    $font = New-Object System.Drawing.Font('Segoe UI',[single]8,[System.Drawing.FontStyle]::Bold)
    $tb = New-Object System.Drawing.SolidBrush ($C_TEXT)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = 'Center'
    $sf.LineAlignment = 'Center'
    $rect = New-Object System.Drawing.RectangleF($x, $y, $w, $h)
    $g.DrawString($txt, $font, $tb, $rect, $sf)
    $font.Dispose(); $tb.Dispose(); $sf.Dispose()
})
$form.Controls.Add($pnlProg)

# 更新進度：百分比 + 顏色（刪除中使用 danger 紅，完成切 OK 綠）
function Set-DeleteProgress {
    param([int]$Done, [int]$Total)
    if ($Total -le 0) { return }
    $script:ProgPct = [int]([math]::Round(100.0 * $Done / $Total))
    $pnlProg.Invalidate()
}

# 固定 Dock 處理順序：Top(pnlTop 全寬) → Top(pnlRibbon 其下) → Bottom → Left → Fill
# WinForms 由「索引最大者先 Dock」，因此重新排列一次確保版面正確
$form.Controls.Clear()
foreach ($ctrl in @($lvResults, $pnlLeft, $pnlBottom, $pnlProg, $status, $pnlRibbon, $pnlTop)) {
    $form.Controls.Add($ctrl)
}

# ---------------------------------------------------------------- 磁碟機列舉
function Update-Drives {
    param([switch]$Silent)
    $clbDrives.Items.Clear()
    $idx = 0
    foreach ($d in [System.IO.DriveInfo]::GetDrives()) {
        if (-not $d.IsReady) { continue }
        $freeGb = [math]::Round($d.AvailableFreeSpace / 1GB, 1)
        $totalGb = [math]::Round($d.TotalSize / 1GB, 1)
        $label = $d.DriveType
        $obj = [pscustomobject]@{
            Root = $d.RootDirectory.FullName
            Text = ('{0}  {1}  ' -f $d.Name, $label) + ((L 'Drive.Available') -f $freeGb, $totalGb)
        }
        [void]$clbDrives.Items.Add($obj)
        if ($d.DriveType -eq 'Fixed' -and $idx -eq 0) {
            $clbDrives.SetItemChecked($clbDrives.Items.Count - 1, $true)
            $idx++
        }
    }
    if (-not $Silent) { Write-Log 'INFO' ((L 'Log.Drives') -f $clbDrives.Items.Count) }
}
Update-Drives

# ---------------------------------------------------------------- 搜尋引擎（背景 runspace）
$script:Searching = $false
$script:State = $null
$script:PS = $null
$script:Handle = $null
$script:AbortAll = $false   # 「立即停止」旗標：一旦設定，所有作業在下一個檢查點停下並結束程式
$script:Deleting = $false   # 刪除迴圈進行中（UI 執行緒），此時由迴圈檢查點負責收尾關閉
$script:LastExt = ''
$script:LastZw = 0

$searchBlock = {
    param($Roots, $Keyword, $TargetIndex, $MaxResults, $MaxDepth, $Zw3dMode, $ExtFilter, $State, $LangWords, $TypeFile, $TypeDir)
    $results = New-Object System.Collections.ArrayList
    $scanned = 0
    $errors = 0
    $kw = [string]$Keyword
    $hasKw = ($kw.Length -gt 0)
    $extF = [string]$ExtFilter
    $hasExt = ($extF.Length -gt 0)
    $wantDir = ($TargetIndex -ne 1)
    $wantFile = ($TargetIndex -ne 0)
    $stop = $false

    # ---- ZW3D 判斷規則（Mode：0=關閉 1=全部 ZW3D 2=僅暫存垃圾 3=僅模型檔）
    $zwExtJunk = @('.z3bak', '.z3gr', '.orig', '.cfg', '.log', '.stl', '.lim', '.ptss', '.job', '.tls',
                   '.wkp', '.projector', '.offsetrep', '.vxbcl', '.info', '.xml', '.bmp')
    $zwExtModel = @('.z3', '.z3prt', '.z3asm', '.z3drw', '.z3sht', '.z3fnt', '.z3cvx')
    $zwNameKeys = @('zw3d', 'topologicrep', 'proxy', 'analyticalrep', 'z3_a7880')
    # 查表改用 hashtable（O(1)），語言關鍵詞編譯成一條正則（取代逐詞 IndexOf 迴圈）
    $junkSet = @{}; foreach ($e in $zwExtJunk) { $junkSet[$e] = $true }
    $modelSet = @{}; foreach ($e in $zwExtModel) { $modelSet[$e] = $true }
    # 語言詞改用單一正則比對（取代逐詞 IndexOf）；不用 Compiled，避免每次搜尋都付 JIT 編譯成本
    $langPattern = $null
    if ($LangWords.Count -gt 0) {
        $langPattern = (($LangWords | ForEach-Object { [regex]::Escape([string]$_) }) -join '|')
    }

    function Test-Zw3dHit {
        param([string]$Name, [string]$Path, [int]$Mode)
        if ($Mode -le 0) { return $true }
        $ext = ''
        $dot = $Name.LastIndexOf('.')
        if ($dot -ge 0) { $ext = $Name.Substring($dot).ToLowerInvariant() }
        $lower = $Name.ToLowerInvariant()
        $pathLower = $Path.ToLowerInvariant()

        # 語言版本關鍵詞：只在已與 ZW3D 相關的名稱／路徑上生效，且排除安裝目錄，避免誤刪系統資源
        if ($langPattern -and
            $pathLower.IndexOf('\program files') -lt 0 -and $pathLower.IndexOf('\languages\') -lt 0) {
            $zwRelated = ($pathLower.IndexOf('zw3d') -ge 0 -or $pathLower.IndexOf('zwsoft') -ge 0 -or
                          $lower.IndexOf('zw3d') -ge 0 -or $junkSet.ContainsKey($ext) -or $modelSet.ContainsKey($ext))
            if ($zwRelated -and [regex]::IsMatch($lower, $langPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) { return $true }
        }

        if ($Mode -eq 2) {
            if ($junkSet.ContainsKey($ext)) { return $true }
            foreach ($k in $zwNameKeys) { if ($lower.IndexOf($k) -ge 0) { return $true } }
            return $false
        }
        if ($Mode -eq 3) {
            if ($modelSet.ContainsKey($ext)) { return $true }
            if ($lower.IndexOf('zw3d') -ge 0) { return $true }
            return $false
        }
        # Mode 1：副檔名 or 檔名關鍵字 or 路徑含 ZW3D
        if ($junkSet.ContainsKey($ext)) { return $true }
        if ($modelSet.ContainsKey($ext)) { return $true }
        foreach ($k in $zwNameKeys) { if ($lower.IndexOf($k) -ge 0) { return $true } }
        if ($pathLower.IndexOf('zw3d') -ge 0) { return $true }
        return $false
    }

    foreach ($root in $Roots) {
        if ($stop) { break }
        $stack = New-Object System.Collections.Stack
        $stack.Push(@($root, 0))
        while ($stack.Count -gt 0) {
            if ($State.Cancel) { $State.Results = $results; $State.Done = $true; return }
            $cur = $stack.Pop()
            $dir = [string]$cur[0]
            $depth = [int]$cur[1]
            $State.Scanning = $dir

            try {
                $subDirs = @([System.IO.Directory]::EnumerateDirectories($dir))
            } catch {
                $errors++
                $State.Errors = $errors
                continue
            }

            foreach ($s in $subDirs) {
                $scanned++
                if ($wantDir) {
                    $nm = [System.IO.Path]::GetFileName($s)
                    if (((-not $hasKw) -or $nm.IndexOf($kw, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) -and
                        (Test-Zw3dHit -Name $nm -Path $s -Mode $Zw3dMode)) {
                        [void]$results.Add([pscustomobject]@{ Name = $nm; Path = $s; Type = $TypeDir; Parent = [System.IO.Path]::GetDirectoryName($s) })
                        if ($results.Count -ge $MaxResults) { $State.Truncated = $true; $stop = $true; break }
                    }
                }
                if ($depth -lt $MaxDepth) { $stack.Push(@($s, $depth + 1)) }
            }

            if ($stop) { break }

            if ($wantFile) {
                try {
                    $files = @([System.IO.Directory]::EnumerateFiles($dir))
                } catch {
                    $errors++
                    $State.Errors = $errors
                    $files = @()
                }
                foreach ($f in $files) {
                    $scanned++
                    $nm = [System.IO.Path]::GetFileName($f)
                    if ((($hasExt -eq $false) -or $nm.EndsWith($extF, [System.StringComparison]::OrdinalIgnoreCase)) -and
                        (((-not $hasKw) -or $nm.IndexOf($kw, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)) -and
                        (Test-Zw3dHit -Name $nm -Path $f -Mode $Zw3dMode)) {
                        [void]$results.Add([pscustomobject]@{ Name = $nm; Path = $f; Type = $TypeFile; Parent = [System.IO.Path]::GetDirectoryName($f) })
                        if ($results.Count -ge $MaxResults) { $State.Truncated = $true; $stop = $true; break }
                    }
                }
            }

            $State.Scanned = $scanned
            $State.Errors = $errors
        }
    }
    $State.Results = $results
    $State.Done = $true
}

$script:Timer = New-Object System.Windows.Forms.Timer
$script:Timer.Interval = 200

function Complete-Search {
    $script:Timer.Stop()
    $s = $script:State
    $out = $null
    try {
        if ($script:PS -and $script:Handle) { $out = $script:PS.EndInvoke($script:Handle) }
    } catch {
        Write-Log 'ERROR' ((L 'Log.SearchThreadErr') -f $_.Exception.Message)
    } finally {
        if ($script:PS) { try { $script:PS.Dispose() } catch { } }
        $script:PS = $null
        $script:Handle = $null
    }

    $results = $null
    if ($s -and $s.Results) { $results = $s.Results }
    if (-not $results) { $results = New-Object System.Collections.ArrayList }

    $items = New-Object 'System.Collections.Generic.List[System.Windows.Forms.ListViewItem]'
    $typeFile = [string](L 'Type.File')
    $typeFolder = [string](L 'Type.Folder')
    foreach ($r in $results) {
        $it = New-Object System.Windows.Forms.ListViewItem([string]$r.Name)
        [void]$it.SubItems.Add($(if ($r.Type -eq 'Folder') { $typeFolder } else { $typeFile }))
        [void]$it.SubItems.Add([string]$r.Path)
        [void]$it.SubItems.Add([string]$r.Parent)
        $it.Tag = [string]$r.Path
        [void]$items.Add($it)
    }
    $lvResults.BeginUpdate()
    try {
        $lvResults.Items.Clear()
        $lvResults.Items.AddRange($items.ToArray())   # 一次大量加入，取代 5000 次 Add
    } finally { $lvResults.EndUpdate() }

    $script:Searching = $false
    $btnSearch.Enabled = $true
    $btnStop.Enabled = $false

    $msg = (L 'Status.Found') -f $lvResults.Items.Count, $s.Scanned, $s.Errors
    if ($s.Truncated) { $msg += L 'Status.Truncated' }
    if ($lvResults.Items.Count -eq 0 -and $script:LastExt -and $script:LastZw -gt 0) {
        $msg += ((L 'Status.ExtHint') -f $script:LastExt)
    }
    $lblStatus.Text = $msg
    $lblScan.Text = ''
    Write-Log 'INFO' ((L 'Log.SearchEnd') -f $s.Keyword, $s.Target, $lvResults.Items.Count, $s.Scanned, $s.Errors)
    $script:State = $null
}

$script:Timer.Add_Tick({
    $s = $script:State
    if (-not $s) { $script:Timer.Stop(); return }
    $lblStatus.Text = (L 'Status.ScanProgress') -f $s.Scanned, $(if ($s.Results) { $s.Results.Count } else { 0 }), $s.Errors
    $cur = [string]$s.Scanning
    if ($cur.Length -gt 90) { $cur = '…' + $cur.Substring($cur.Length - 90) }
    $lblScan.Text = $cur
    if ($s.Done) { Complete-Search }
})

# 取得目前選擇的副檔名（正規化：小寫、補上開頭的點）；未選擇回傳空字串
function Get-ExtFilter {
    $v = [string]$cboExt.Text
    $v = $v.Trim()
    if ($cboTarget.SelectedIndex -eq 0) { return '' }      # 資料夾模式不套副檔名
    if ($v -eq '' -or $v -eq (L 'Ext.All') -or $v -in @('*', '.*', '不限', 'all')) { return '' }
    if (-not $v.StartsWith('.')) { $v = '.' + $v }
    return $v.ToLowerInvariant()
}

$btnSearch.Add_Click({
    if ($script:Searching) { return }
    $roots = @()
    foreach ($i in $clbDrives.CheckedItems) { $roots += [string]$i.Root }
    if ($roots.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show((L 'Msg.NoDrive'), (L 'Msg.InfoTitle'), 'OK', 'Information') | Out-Null
        return
    }
    $kw = $txtKeyword.Text.Trim()
    if ($cboPreset.SelectedIndex -gt 0 -and $kw.Length -eq 0) { $kw = [string]$cboPreset.SelectedItem }
    if ($kw -eq (L 'Preset.All')) { $kw = '' }

    # 副檔名只套用在檔案；搜尋對象＝資料夾時 Get-ExtFilter 回傳空字串（不再強制切到「檔案」）
    $extFilter = Get-ExtFilter

    $depth = 12
    switch ($cboDepth.SelectedIndex) {
        0 { $depth = 3 } 1 { $depth = 5 } 2 { $depth = 8 } 3 { $depth = 12 } 4 { $depth = 999 }
    }

    # ZW3D 篩選模式：0=關閉 1=全部 2=僅暫存垃圾 3=僅模型檔
    $zwMode = 0
    if ($chkZw3dOnly.Checked) { $zwMode = $cboZw3dScope.SelectedIndex + 1 }
    $zwText = L 'ZwMode.Off'
    if ($zwMode -eq 1) { $zwText = L 'ZwMode.All' }
    elseif ($zwMode -eq 2) { $zwText = L 'ZwMode.Junk' }
    elseif ($zwMode -eq 3) { $zwText = L 'ZwMode.Model' }

    $lvResults.Items.Clear()
    $script:State = [hashtable]::Synchronized(@{
        Done = $false; Cancel = $false; Scanning = ''; Scanned = 0; Errors = 0
        Results = $null; Truncated = $false; Keyword = $kw; Target = [string]$cboTarget.SelectedItem
    })
    $script:Searching = $true
    $btnSearch.Enabled = $false
    $btnStop.Enabled = $true
    $lblStatus.Text = L 'Status.Scanning'
    $script:LastExt = $extFilter
    $script:LastZw = $zwMode
    Write-Log 'INFO' ((L 'Log.SearchStart') -f ($roots -join ', '), $kw, $cboTarget.SelectedItem, $depth, $zwText, $(if ($extFilter) { $extFilter } else { L 'Ext.Unlimited' }), $(if ($script:ScanLangId) { $script:ScanLangId } else { 'All' }))

    try {
        $script:PS = [powershell]::Create()
        $script:PS.Runspace = [runspacefactory]::CreateRunspace()
        $script:PS.Runspace.ApartmentState = 'STA'
        $script:PS.Runspace.Open()
        [void]$script:PS.AddScript($searchBlock)
        [void]$script:PS.AddArgument($roots)
        [void]$script:PS.AddArgument($kw)
        [void]$script:PS.AddArgument([int]$cboTarget.SelectedIndex)
        [void]$script:PS.AddArgument(5000)
        [void]$script:PS.AddArgument($depth)
        [void]$script:PS.AddArgument($zwMode)
        [void]$script:PS.AddArgument($extFilter)
        [void]$script:PS.AddArgument($script:State)
        [void]$script:PS.AddArgument(@(Get-LangWords $script:ScanLangId))
        [void]$script:PS.AddArgument([string](L 'Type.File'))
        [void]$script:PS.AddArgument([string](L 'Type.Folder'))
        $script:Handle = $script:PS.BeginInvoke()
        $script:Timer.Start()
    } catch {
        Write-Log 'ERROR' ((L 'Log.SearchStartFail') -f $_.Exception.Message)
        $lblStatus.Text = L 'Status.SearchFail'
        $script:Searching = $false
        $btnSearch.Enabled = $true
        $btnStop.Enabled = $false
    }
})

$btnStop.Add_Click({
    if ($script:State) { $script:State.Cancel = $true; Write-Log 'WARN' (L 'Log.UserStop') }
})

# ---------------------------------------------------------------- 立即停止（中止所有作業並結束程式）
# 觸發：Ribbon「立即停止」按鈕，或快速鍵 Ctrl+Q（表單 KeyPreview）。
# 行為：設定 AbortAll 旗標 → 停掉進度 Timer → 取消並停止背景掃描 runspace →
#       閒置/掃描中直接關閉主視窗；刪除迴圈進行中則由迴圈檢查點收尾後關閉。
function Invoke-AbortAll {
    if ($script:AbortAll) { return }   # 防止重複觸發
    $script:AbortAll = $true
    Write-Log 'WARN' (L 'Log.Abort')
    try { $lblStatus.Text = L 'Status.Aborting' } catch { }
    try { $script:Timer.Stop() } catch { }
    if ($script:State) { $script:State.Cancel = $true }
    if ($script:PS) {
        try { $script:PS.Stop() } catch { }
        try { $script:PS.Dispose() } catch { }
        $script:PS = $null
    }
    if (-not $script:Deleting) {
        try { $form.Close() } catch { }   # FormClosing 會再做一次等冪清理
    }
    # 刪除迴圈進行中：此處正在 DoEvents 內，不能關視窗；由迴圈檢查點偵測旗標後收尾
}
$btnAbort.Add_Click({ Invoke-AbortAll })

$txtKeyword.Add_KeyDown({
    if ($_.KeyCode -eq 'Return') { $btnSearch.PerformClick() }
})
$txtKeyword.Add_TextChanged({
    if ($cboPreset.SelectedIndex -gt 0) { $cboPreset.SelectedIndex = 0 }
})
$cboPreset.Add_SelectedIndexChanged({
    if ($script:ApplyingLang) { return }
    if ($cboPreset.SelectedIndex -gt 0) { $txtKeyword.Text = [string]$cboPreset.SelectedItem }
})

# ---------------------------------------------------------------- 勾選 / 全選
$script:CheckAllAction = {
    $lvResults.BeginUpdate()
    foreach ($it in $lvResults.Items) { $it.Checked = $true }
    $lvResults.EndUpdate()
    $lblStatus.Text = (L 'Status.Checked') -f $lvResults.CheckedItems.Count
}
$btnCheckResults.Add_Click($script:CheckAllAction)
$btnUncheckResults.Add_Click({
    $lvResults.BeginUpdate()
    foreach ($it in $lvResults.Items) { $it.Checked = $false }
    $lvResults.EndUpdate()
    $lblStatus.Text = L 'Status.Unchecked'
})
$btnCheckAll.Add_Click({
    $all = $true
    for ($i = 0; $i -lt $clbDrives.Items.Count; $i++) {
        if (-not $clbDrives.GetItemChecked($i)) { $all = $false; break }
    }
    for ($i = 0; $i -lt $clbDrives.Items.Count; $i++) {
        $clbDrives.SetItemChecked($i, -not $all)
    }
})
$btnRefresh.Add_Click({ Update-Drives; $lblStatus.Text = L 'Status.DrivesRefreshed' })
$btnOpenLog.Add_Click({
    try { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $script:LogDir) } catch { }
})

# 匯出目前清單：副檔名依下拉選單自動套用（僅支援 .txt / .csv / .json，其餘回退 .csv）
$btnExport.Add_Click({
    if ($lvResults.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show((L 'Msg.NoExport'), (L 'Msg.InfoTitle'), 'OK', 'Information') | Out-Null
        return
    }

    $ext = Get-ExtFilter
    if ($script:ExportExts -notcontains $ext) { $ext = '.csv' }

    $rows = @()
    foreach ($it in $lvResults.Items) {
        $rows += [pscustomobject]@{
            名稱 = $it.SubItems[0].Text
            類型 = $it.SubItems[1].Text
            路徑 = $it.SubItems[2].Text
            上層 = $it.SubItems[3].Text
        }
    }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title = L 'Export.Title'
    switch ($ext) {
        '.txt'  { $dlg.Filter = L 'Export.FilterTxt' }
        '.json' { $dlg.Filter = L 'Export.FilterJson' }
        default { $dlg.Filter = L 'Export.FilterCsv' }
    }
    $dlg.DefaultExt = $ext.TrimStart('.')
    $dlg.FileName = ('{0}{1}{2}' -f (L 'Export.FileName'), (Get-Date -Format 'yyyyMMdd_HHmm'), $ext)
    $dlg.InitialDirectory = $script:BaseDir
    if ($dlg.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $outPath = $dlg.FileName
    $dlg.Dispose()

    $content = ''
    switch ($ext) {
        '.txt'  { $content = (($rows | ForEach-Object { $_.路徑 }) -join "`r`n") }
        '.json' { $content = ($rows | ConvertTo-Json -Depth 4) }
        default { $content = ($rows | ConvertTo-Csv -NoTypeInformation | Out-String) }
    }

    try {
        [System.IO.File]::WriteAllText($outPath, $content, (New-Object System.Text.UTF8Encoding($true)))
        Write-Log 'INFO' ((L 'Log.ExportOk') -f $ext, $outPath, $rows.Count)
        $lblStatus.Text = ((L 'Status.Exported') -f $rows.Count, $outPath)
        [System.Windows.Forms.MessageBox]::Show(((L 'Msg.ExportOk') -f $rows.Count, $outPath), (L 'Msg.ExportOkTitle'), 'OK', 'Information') | Out-Null
    } catch {
        Write-Log 'FAIL' ((L 'Log.ExportFail') -f $outPath, $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show(((L 'Msg.ExportFailBody') + $_.Exception.Message), (L 'Msg.ErrTitle'), 'OK', 'Error') | Out-Null
    }
})
$btnOpenFolder.Add_Click({
    if ($lvResults.SelectedItems.Count -gt 0) {
        $p = [string]$lvResults.SelectedItems[0].Tag
        try {
            if (Test-Path -LiteralPath $p -PathType Container) {
                Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $p)
            } else {
                Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"{0}"' -f $p)
            }
        } catch { Write-Log 'ERROR' ((L 'Log.OpenFail') -f $p) }
    }
})

# 右鍵選單
$menu = New-Object System.Windows.Forms.ContextMenuStrip
$menu.BackColor = $C_FIELD
$menu.ForeColor = $C_TEXT
$miOpen = $menu.Items.Add((L 'Menu.Open'))
$miCopy = $menu.Items.Add((L 'Menu.Copy'))
$miDelete = $menu.Items.Add((L 'Menu.Delete'))
$miOpen.Add_Click({ $btnOpenFolder.PerformClick() })
$miCopy.Add_Click({
    if ($lvResults.SelectedItems.Count -gt 0) {
        try { [System.Windows.Forms.Clipboard]::SetText([string]$lvResults.SelectedItems[0].Tag) } catch { }
    }
})
$miDelete.Add_Click({
    foreach ($it in $lvResults.SelectedItems) { $it.Checked = $true }
    $btnDelete.PerformClick()
})
$lvResults.ContextMenuStrip = $menu

# ---------------------------------------------------------------- 語言套用
# 靜態控制項文案集中註冊，切換語言時統一重刷；動態文案（日誌／狀態／彈窗）直接呼叫 L()。
$script:UiBindings = @(
    @{ C = $form;          K = 'AppTitle' }
    @{ C = $lblKw;         K = 'KwLabel' }
    @{ C = $lblPreset;     K = 'PresetLabel' }
    @{ C = $lblTarget;     K = 'TargetLabel' }
    @{ C = $lblDepth;      K = 'DepthLabel' }
    @{ C = $chkPermanent;  K = 'PermanentChk' }
    @{ C = $lblHint;       K = 'Hint' }
    @{ C = $chkZw3dOnly;   K = 'Zw3dOnlyChk' }
    @{ C = $lblExt;        K = 'ExtLabel' }
    @{ C = $lblZw3dHint;   K = 'Zw3dHint' }
    @{ C = $lblUi;         K = 'UiLangLabel' }
    @{ C = $lblLang;       K = 'LangLabel' }
    @{ C = $lblDrives;     K = 'DrivesLabel' }
    @{ C = $lblStatus;     K = 'Status.Ready' }
    @{ C = $miOpen;        K = 'Menu.Open' }
    @{ C = $miCopy;        K = 'Menu.Copy' }
    @{ C = $miDelete;      K = 'Menu.Delete' }
)

function Set-ComboKeys($cbo, [string[]]$keys) {
    $idx = $cbo.SelectedIndex
    $txt = [string]$cbo.Text
    $editable = ($cbo.DropDownStyle -ne 'DropDownList')
    $cbo.BeginUpdate()
    $cbo.Items.Clear()
    foreach ($k in $keys) { [void]$cbo.Items.Add((L $k)) }
    $cbo.EndUpdate()
    if ($idx -ge 0 -and $idx -lt $cbo.Items.Count) { $cbo.SelectedIndex = $idx }
    elseif ($editable) { $cbo.Text = $txt }
    elseif ($cbo.Items.Count -gt 0) { $cbo.SelectedIndex = 0 }
}

function Update-UiFonts($parent) {
    $fam = Get-UiFontFamily
    foreach ($c in $parent.Controls) {
        try {
            if ($c.Font -and @('Consolas', 'Segoe UI') -notcontains $c.Font.Name) {
                $c.Font = New-Object System.Drawing.Font($fam, $c.Font.Size, $c.Font.Style)
            }
        } catch { }
        if ($c.Controls.Count -gt 0) { Update-UiFonts $c }
    }
}

function Set-Language {
    $script:ApplyingLang = $true
    try {
        foreach ($b in $script:UiBindings) { $b.C.Text = L $b.K }

        $lvResults.Columns[0].Text = L 'Col.Name'
        $lvResults.Columns[1].Text = L 'Col.Type'
        $lvResults.Columns[2].Text = L 'Col.Path'
        $lvResults.Columns[3].Text = L 'Col.Parent'
        $lblCredit.ToolTipText = L 'CreditTip'

        foreach ($b in $script:IconButtons) {
            $b.CapLabel.Text = L $b.CapKey
            $tip = L $b.TipKey
            $script:Tip.SetToolTip($b, $tip)
            foreach ($child in $b.Controls) { $script:Tip.SetToolTip($child, $tip) }
        }
        $script:Tip.SetToolTip($cboUi, (L 'Tip.UiLang'))
        $script:Tip.SetToolTip($lblUi, (L 'Tip.UiLang'))
        $script:Tip.SetToolTip($cboLang, (L 'Tip.ScanLang'))
        $script:Tip.SetToolTip($lblLang, (L 'Tip.ScanLang'))

        Set-ComboKeys $cboPreset (@('Preset.All') + @(Get-PresetWords $script:ScanLangId))
        Set-ComboKeys $cboTarget @('Target.Folder', 'Target.File', 'Target.Both')
        Set-ComboKeys $cboDepth @('Depth.3', 'Depth.5', 'Depth.8', 'Depth.12', 'Depth.Unlimited')
        Set-ComboKeys $cboZw3dScope @('ZwScope.All', 'ZwScope.Junk', 'ZwScope.Model')
        Set-ComboKeys $cboExt (@('Ext.All') + @($script:ExtPresets | Select-Object -Skip 1))

        # 掃描語言下拉的第 0 項「全部語言」也跟著介面語言（其餘為語言自稱，不翻譯）
        $scanIdx = $cboLang.SelectedIndex
        $cboLang.BeginUpdate()
        $cboLang.Items[0] = L 'Lang.All'
        $cboLang.EndUpdate()
        if ($cboLang.SelectedIndex -ne $scanIdx) { $cboLang.SelectedIndex = $scanIdx }

        # 磁碟機列文字含「GB 可用」，一併重刷（保留勾選狀態）
        $checkedRoots = @()
        foreach ($i in $clbDrives.CheckedItems) { $checkedRoots += [string]$i.Root }
        Update-Drives -Silent
        for ($i = 0; $i -lt $clbDrives.Items.Count; $i++) {
            if ($checkedRoots -contains [string]$clbDrives.Items[$i].Root) { $clbDrives.SetItemChecked($i, $true) }
        }

        $fam = Get-UiFontFamily
        try { $form.Font = New-Object System.Drawing.Font($fam, 10) } catch { }
        Update-UiFonts $form
    } finally {
        $script:ApplyingLang = $false
    }
}

# 介面語言下拉：只改介面文案與字體
$cboUi.Add_SelectedIndexChanged({
    if ($script:ApplyingLang) { return }
    $i = $cboUi.SelectedIndex
    if ($i -lt 0 -or $i -ge $script:Langs.Count) { return }
    $script:UiLang = $script:Langs[$i].Id
    Save-LangPref
    Set-Language
    Write-Log 'INFO' ((L 'Log.UiLang') -f $script:UiLang)
})

# ZW3D 語言版本下拉：只改掃描關鍵詞與常用預設，介面文案不變
$cboLang.Add_SelectedIndexChanged({
    if ($script:ApplyingLang) { return }
    $i = $cboLang.SelectedIndex
    if ($i -le 0) { $script:ScanLangId = '' }
    else { $script:ScanLangId = $script:Langs[$i - 1].Id }
    Save-LangPref
    $script:ApplyingLang = $true
    try { Set-ComboKeys $cboPreset (@('Preset.All') + @(Get-PresetWords $script:ScanLangId)) }
    finally { $script:ApplyingLang = $false }
    Write-Log 'INFO' ((L 'Log.ScanLang') -f $(if ($script:ScanLangId) { $script:ScanLangId } else { 'All' }))
})

Set-Language

# ---------------------------------------------------------------- 刪除
function Test-ProtectedPath {
    param([string]$Path)
    $p = $Path.TrimEnd('\')
    if ($p -match '^[A-Za-z]:$') { return $true }
    $guards = @($env:SystemRoot, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramData, $env:USERPROFILE, "$env:SystemDrive\Users", "$env:SystemDrive\Windows", "$env:SystemDrive")
    foreach ($g in $guards) {
        if ($g -and $p -eq $g.TrimEnd('\')) { return $true }
    }
    return $false
}

function Show-DeleteConfirm {
    param([string[]]$Paths, [bool]$Permanent)
    if ($script:PerfMode) { return 'ok' }   # -perf：略過確認視窗（僅效能基準用）

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = L 'Dlg.Title'
    $dlg.Size = New-Object System.Drawing.Size(860, 560)
    $dlg.StartPosition = 'CenterParent'
    $dlg.BackColor = $C_BG
    $dlg.ForeColor = $C_TEXT
    $dlg.Font = $FONT_UI
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $title = New-Object System.Windows.Forms.Label
    $title.Text = (L 'Dlg.Heading') -f $Paths.Count
    $title.Font = New-Object System.Drawing.Font('Microsoft JhengHei UI', 14, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = $C_WARN
    $title.AutoSize = $true
    $title.Location = New-Object System.Drawing.Point(16, 14)
    $dlg.Controls.Add($title)

    $desc = New-Object System.Windows.Forms.Label
    if ($Permanent) {
        $desc.Text = L 'Dlg.Perm'
        $desc.ForeColor = $C_DANGER
    } else {
        $desc.Text = L 'Dlg.Recycle'
        $desc.ForeColor = $C_OK
    }
    $desc.AutoSize = $true
    $desc.Location = New-Object System.Drawing.Point(18, 52)
    $dlg.Controls.Add($desc)

    $list = New-Object System.Windows.Forms.ListBox
    $list.Location = New-Object System.Drawing.Point(18, 80)
    $list.Size = New-Object System.Drawing.Size(806, 330)
    $list.BackColor = [System.Drawing.Color]::FromArgb(22, 22, 24)
    $list.ForeColor = $C_TEXT
    $list.Font = $FONT_MONO
    $list.BorderStyle = 'FixedSingle'
    $list.HorizontalScrollbar = $true
    foreach ($p in $Paths) { [void]$list.Items.Add($p) }
    $dlg.Controls.Add($list)

    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = L 'Dlg.Confirm'
    $chk.AutoSize = $true
    $chk.Location = New-Object System.Drawing.Point(20, 424)
    $chk.ForeColor = $C_TEXT
    $dlg.Controls.Add($chk)

    $btnCancel = New-DarkButton -Text (L 'Dlg.Cancel') -Width 110 -Height 36
    $btnCancel.Location = New-Object System.Drawing.Point(590, 462)
    $dlg.Controls.Add($btnCancel)

    $btnOk = New-DarkButton -Text (L 'Dlg.Ok') -Width 130 -Height 36 -BackColor $C_DANGER -ForeColor ([System.Drawing.Color]::White)
    $btnOk.Location = New-Object System.Drawing.Point(710, 462)
    $btnOk.Enabled = $false
    $dlg.Controls.Add($btnOk)

    $script:DlgResult = 'cancel'
    $chk.Add_CheckedChanged({ $btnOk.Enabled = $chk.Checked })
    $btnOk.Add_Click({ $script:DlgResult = 'ok'; $dlg.Close() })
    $btnCancel.Add_Click({ $script:DlgResult = 'cancel'; $dlg.Close() })
    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel

    [void]$dlg.ShowDialog($form)
    $r = $script:DlgResult
    $dlg.Dispose()
    return $r
}

function Remove-OneItem {
    param([string]$Path, [bool]$Permanent)
    $isDir = $false
    try { $isDir = [System.IO.Directory]::Exists($Path) } catch { }
    if ((-not $Permanent) -and (-not $script:RecycleOk)) {
        return [pscustomobject]@{ Ok = $false; Reason = (L 'Msg.RecycleFail') }
    }
    try {
        if ($isDir) {
            if ($Permanent) {
                [System.IO.Directory]::Delete($Path, $true)
            } else {
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($Path, 'OnlyErrorDialogs', 'SendToRecycleBin')
            }
        } else {
            if ($Permanent) {
                [System.IO.File]::Delete($Path)
            } else {
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($Path, 'OnlyErrorDialogs', 'SendToRecycleBin')
            }
        }
        return [pscustomobject]@{ Ok = $true; Reason = '' }
    } catch [System.UnauthorizedAccessException] {
        return [pscustomobject]@{ Ok = $false; Reason = (L 'Msg.ErrPermission') }
    } catch [System.Security.SecurityException] {
        return [pscustomobject]@{ Ok = $false; Reason = (L 'Msg.ErrSecurity') }
    } catch [System.IO.IOException] {
        $m = $_.Exception.Message
        if ($m -match '使用中|in use|being used|另一個程序|locked') {
            return [pscustomobject]@{ Ok = $false; Reason = (L 'Msg.ErrLocked') + $m }
        }
        if ($m -match 'too long|過長|path') {
            return [pscustomobject]@{ Ok = $false; Reason = (L 'Msg.ErrLongPath') + $m }
        }
        return [pscustomobject]@{ Ok = $false; Reason = (L 'Msg.ErrIO') + $m }
    } catch {
        return [pscustomobject]@{ Ok = $false; Reason = ((L 'Msg.ErrUnexpected') -f $_.Exception.Message) }
    }
}


# ---------------------------------------------------------------- ZW3D 使用者目錄清理
# 對應官方 FolderMgr.exe 的四個動作。以「實例目錄」為單位（%APPDATA%\ZWSOFT\ZW3D\<版本>）。
# 清理採整目錄刪除：output / custom / temp / 整個實例；刪除前一律預覽 + 二次確認。
function Get-UserFolderInstances([string]$Root) {
    if ([string]::IsNullOrEmpty($Root)) { $Root = Join-Path $env:APPDATA 'ZWSOFT\ZW3D' }
    if (-not (Test-Path -LiteralPath $Root)) { return @() }
    $out = New-Object System.Collections.ArrayList
    foreach ($d in [System.IO.Directory]::EnumerateDirectories($Root)) {
        if ((Test-Path -LiteralPath (Join-Path $d 'custom')) -or
            (Test-Path -LiteralPath (Join-Path $d 'output')) -or
            (Test-Path -LiteralPath (Join-Path $d 'temp'))) { [void]$out.Add($d) }
    }
    return @($out.ToArray() | Sort-Object -Descending)
}

# 版本下拉用的顯示名 ↔ 完整路徑。實例偵測沿用 Get-UserFolderInstances（本機實測：
# %APPDATA%\ZWSOFT\ZW3D\ 下僅有含 custom/output/temp 的目錄才是真正的版本使用者目錄）。
function Get-UfVersionItems([string]$Root) {
    $out = New-Object System.Collections.ArrayList
    foreach ($p in @(Get-UserFolderInstances $Root)) {
        [void]$out.Add([pscustomobject]@{ Name = (Split-Path -Leaf $p); Path = $p })
    }
    return $out.ToArray()
}

function Get-UserCleanTarget([string]$Instance, [string]$Kind) {
    switch ($Kind) {
        'output' { return (Join-Path $Instance 'output') }
        'custom' { return (Join-Path $Instance 'custom') }
        'temp'   { return (Join-Path $Instance 'temp') }
        'all'    { return $Instance }
    }
    throw ('unknown clean kind: {0}' -f $Kind)
}

function Get-UserCleanStats([string]$Target) {
    $files = 0; $dirs = 0
    try { foreach ($f in [System.IO.Directory]::EnumerateFiles($Target, '*', 'AllDirectories')) { $files++ } } catch { }
    try { foreach ($d in [System.IO.Directory]::EnumerateDirectories($Target, '*', 'AllDirectories')) { $dirs++ } } catch { }
    return [pscustomobject]@{ Files = $files; Dirs = $dirs }
}

function Format-UfSize([long]$Bytes) {
    if ($Bytes -lt 1024) { return ('{0} B' -f $Bytes) }
    if ($Bytes -lt 1048576) { return ('{0:N1} KB' -f ($Bytes / 1024)) }
    if ($Bytes -lt 1073741824) { return ('{0:N1} MB' -f ($Bytes / 1048576)) }
    return ('{0:N2} GB' -f ($Bytes / 1073741824))
}

# 一次遍歷整個版本目錄，同時算出 all 與 output/custom/temp 的檔案數、資料夾數、總大小。
# 有 30000 檔 / 800ms 雙重上限，避免切換版本時卡住介面。
function Get-UfInstanceStats([string]$Instance) {
    $r = @{}
    foreach ($k in @('output', 'custom', 'temp', 'all')) {
        $r[$k] = @{ Files = 0; Dirs = 0; Bytes = [long]0 }
    }
    $r['Truncated'] = $false
    if (-not (Test-Path -LiteralPath $Instance)) { return $r }
    $all = $r['all']
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $n = 0
    try {
        $di = [System.IO.DirectoryInfo]::new($Instance)
        foreach ($fi in $di.EnumerateFiles('*', [System.IO.SearchOption]::AllDirectories)) {
            $n++
            if ($n -gt 30000 -or $sw.ElapsedMilliseconds -gt 800) { $r['Truncated'] = $true; break }
            $all.Files++
            $all.Bytes += [long]$fi.Length
            $rel = $fi.FullName.Substring($Instance.Length).TrimStart('\')
            foreach ($s in @('output', 'custom', 'temp')) {
                if ($rel.StartsWith(($s + '\'), [System.StringComparison]::OrdinalIgnoreCase)) {
                    $one = $r[$s]; $one.Files++; $one.Bytes += [long]$fi.Length; break
                }
            }
        }
        foreach ($sd in $di.EnumerateDirectories('*', [System.IO.SearchOption]::AllDirectories)) {
            $all.Dirs++
            $rel = $sd.FullName.Substring($Instance.Length).TrimStart('\')
            foreach ($s in @('output', 'custom', 'temp')) {
                if ($rel.StartsWith(($s + '\'), [System.StringComparison]::OrdinalIgnoreCase)) {
                    $one = $r[$s]; $one.Dirs++; break
                }
            }
        }
    } catch { }
    return $r
}

function Format-UfStat($Stat, [bool]$Exists) {
    if (-not $Exists) { return L 'Uf.NoTarget' }
    return ((L 'Uf.StatLine') -f $Stat.Files, $Stat.Dirs, (Format-UfSize ([long]$Stat.Bytes)))
}

# 切換版本後，位置與四個選項的目標路徑、統計一次全部更新
function Update-UfSelection {
    if ($script:UfRefreshing) { return }
    if (-not $script:UfCbo -or $script:UfCbo.IsDisposed) { return }
    $idx = $script:UfCbo.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:UfPaths.Count) {
        $script:UfInstance = ''
        $script:UfLoc.Text = ''
        foreach ($k in @('output', 'custom', 'temp', 'all')) {
            $script:UfOptPath[$k].Text = ''
            $script:UfOptStat[$k].Text = L 'Uf.NoTarget'
            $script:UfOptBtn[$k].Enabled = $false
        }
        return
    }
    $inst = [string]$script:UfPaths[$idx]
    $script:UfInstance = $inst
    $script:UfLoc.Text = $inst
    $stats = Get-UfInstanceStats $inst
    foreach ($k in @('output', 'custom', 'temp', 'all')) {
        $t = Get-UserCleanTarget $inst $k
        $exists = Test-Path -LiteralPath $t
        $script:UfOptPath[$k].Text = $t
        $script:UfTip.SetToolTip($script:UfOptPath[$k], $t)
        $script:UfOptBtn[$k].Enabled = $exists
        $line = Format-UfStat $stats[$k] $exists
        if ($stats['Truncated']) { $line += ' +' }
        $script:UfOptStat[$k].Text = $line
    }
}

function Refresh-UfInstances([string]$Root, [string]$SelectPath) {
    if (-not $script:UfCbo -or $script:UfCbo.IsDisposed) { return }
    $items = @(Get-UfVersionItems $Root)
    $script:UfRefreshing = $true
    try {
        $script:UfCbo.Items.Clear()
        $script:UfPaths = @($items | ForEach-Object { $_.Path })
        foreach ($it in $items) { [void]$script:UfCbo.Items.Add($it.Name) }
    } finally {
        $script:UfRefreshing = $false
    }
    $idx = 0
    if ($SelectPath) {
        for ($i = 0; $i -lt $script:UfPaths.Count; $i++) {
            if ([string]$script:UfPaths[$i] -ieq $SelectPath) { $idx = $i; break }
        }
    }
    if ($script:UfCbo.Items.Count -gt 0) { $script:UfCbo.SelectedIndex = $idx }
    else { Update-UfSelection }
}

function Invoke-UserClean([string]$Instance, [string]$Kind, [bool]$Permanent) {
    $target = Get-UserCleanTarget $Instance $Kind
    # 護欄：目標必須落在選定的實例目錄內，绝不逸出
    if (-not $target.StartsWith($Instance, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw ('target outside instance: {0}' -f $target)
    }
    if (-not (Test-Path -LiteralPath $target)) {
        return [pscustomobject]@{ Ok = $true; Skipped = $true; Reason = 'missing'; Target = $target }
    }
    $r = Remove-OneItem -Path $target -Permanent $Permanent
    return [pscustomobject]@{ Ok = $r.Ok; Skipped = $false; Reason = $r.Reason; Target = $target }
}

function Invoke-UfAction([string]$Kind, [string]$Instance) {
    $capKey = @{ output='Uf.ClearOutput'; custom='Uf.ClearCustom'; temp='Uf.ClearTemp'; all='Uf.ClearAll' }[$Kind]
    $target = Get-UserCleanTarget $Instance $Kind
    if (-not (Test-Path -LiteralPath $target)) {
        $script:UfStatus.Text = (L 'Uf.Missing') -f $target
        return
    }
    $zwProcs = Get-Process -Name 'zw3d' -ErrorAction SilentlyContinue
    if ($zwProcs) {
        $a = [System.Windows.Forms.MessageBox]::Show((L 'Uf.ZwRunning'), (L 'Uf.ConfirmTitle'), 'YesNo', 'Warning')
        if ($a -ne 'Yes') { $script:UfStatus.Text = L 'Uf.Cancelled'; return }
    }
    $stats = Get-UserCleanStats $target
    $warn = ''
    if ($Kind -eq 'custom') { $warn = L 'Uf.WarnCustom' }
    if ($Kind -eq 'all')    { $warn = L 'Uf.WarnAll' }
    $mode = if ($chkPermanent.Checked) { L 'Mode.Permanent' } else { L 'Mode.Recycle' }
    $msg = ((L 'Uf.Confirm') -f (L $capKey), $target, $stats.Files, $stats.Dirs, $mode, (Split-Path -Leaf $Instance)) + $warn
    $ans = [System.Windows.Forms.MessageBox]::Show($msg, (L 'Uf.ConfirmTitle'), 'YesNo', 'Warning')
    if ($ans -ne 'Yes') { $script:UfStatus.Text = L 'Uf.Cancelled'; return }

    $res = Invoke-UserClean -Instance $Instance -Kind $Kind -Permanent $chkPermanent.Checked
    $resultText = if ($res.Ok) { 'OK' } else { ('FAIL: ' + $res.Reason) }
    Write-Log 'INFO' ((L 'Log.UserClean') -f (L $capKey), $target, $resultText)
    if ($res.Ok) { $script:UfStatus.Text = (L 'Uf.Ok') -f (L $capKey) }
    else { $script:UfStatus.Text = (L 'Uf.Fail') -f $res.Reason }
    # 清理後重新偵測版本（整目錄被刪時該版本會消失），並盡量留在原版本
    Refresh-UfInstances (Split-Path -Parent $Instance) $Instance
}

function Show-UserFolderCleaner {
    param([string]$Root, [switch]$NoShow)

    $items = @(Get-UfVersionItems $Root)
    if ($items.Count -eq 0) {
        if (-not $NoShow) {
            $shownRoot = if ($Root) { $Root } else { (Join-Path $env:APPDATA 'ZWSOFT\ZW3D') }
            [System.Windows.Forms.MessageBox]::Show(
                ((L 'Uf.NoFolder') -f $shownRoot),
                (L 'Dlg.UfTitle'), 'OK', 'Information') | Out-Null
        }
        return $null
    }

    $fam = Get-UiFontFamily
    $smallFont = New-Object System.Drawing.Font($fam, 8.5)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = L 'Dlg.UfTitle'
    $dlg.Size = New-Object System.Drawing.Size(820, 565)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.BackColor = $C_BG
    $dlg.ForeColor = $C_TEXT
    $dlg.Font = New-Object System.Drawing.Font($fam, 10)
    $script:UfDlg = $dlg

    $script:UfTip = New-Object System.Windows.Forms.ToolTip

    # 第一列：選擇要刪除的版本（自動偵測，不需輸入路徑）
    $lblVer = New-Object System.Windows.Forms.Label
    $lblVer.Text = L 'Dlg.UfVersion'
    $lblVer.AutoSize = $true
    $lblVer.Location = New-Object System.Drawing.Point(20, 26)
    $lblVer.ForeColor = $C_TEXT
    $dlg.Controls.Add($lblVer)
    $script:UfVerLabel = $lblVer

    $cboVer = New-Object System.Windows.Forms.ComboBox
    $cboVer.DropDownStyle = 'DropDownList'
    $cboVer.Location = New-Object System.Drawing.Point(230, 22)
    $cboVer.Size = New-Object System.Drawing.Size(360, 28)
    $cboVer.BackColor = $C_FIELD
    $cboVer.ForeColor = $C_TEXT
    $cboVer.FlatStyle = 'Flat'
    $cboVer.Font = $dlg.Font
    $dlg.Controls.Add($cboVer)
    $script:UfCbo = $cboVer
    $script:UfTip.SetToolTip($cboVer, (L 'Dlg.UfHint'))

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = L 'Dlg.UfHint'
    $lblHint.AutoSize = $false
    $lblHint.Location = New-Object System.Drawing.Point(20, 55)
    $lblHint.Size = New-Object System.Drawing.Size(780, 18)
    $lblHint.AutoEllipsis = $true
    $lblHint.Font = $smallFont
    $lblHint.ForeColor = $C_MUTED
    $dlg.Controls.Add($lblHint)
    $script:UfHint = $lblHint
    $script:UfTip.SetToolTip($lblHint, (L 'Dlg.UfHint'))

    # 第二列：位置（唯讀，隨版本自動更新）
    $lblLoc = New-Object System.Windows.Forms.Label
    $lblLoc.Text = L 'Dlg.UfLocation'
    $lblLoc.AutoSize = $true
    $lblLoc.Location = New-Object System.Drawing.Point(20, 88)
    $lblLoc.ForeColor = $C_MUTED
    $dlg.Controls.Add($lblLoc)

    $txtLoc = New-Object System.Windows.Forms.TextBox
    $txtLoc.Location = New-Object System.Drawing.Point(230, 84)
    $txtLoc.Size = New-Object System.Drawing.Size(490, 26)
    $txtLoc.ReadOnly = $true
    $txtLoc.TabStop = $false
    $txtLoc.BackColor = $C_FIELD
    $txtLoc.ForeColor = $C_TEXT
    $txtLoc.BorderStyle = 'FixedSingle'
    $txtLoc.Font = $smallFont
    $txtLoc.WordWrap = $false
    $dlg.Controls.Add($txtLoc)
    $script:UfLoc = $txtLoc
    $script:UfTip.SetToolTip($txtLoc, (L 'Dlg.UfHint'))

    $btnOpen = New-Object System.Windows.Forms.Button
    $btnOpen.Text = L 'Dlg.UfOpen'
    $btnOpen.Location = New-Object System.Drawing.Point(728, 83)
    $btnOpen.Size = New-Object System.Drawing.Size(64, 27)
    $btnOpen.FlatStyle = 'Flat'
    $btnOpen.BackColor = $C_FIELD
    $btnOpen.ForeColor = $C_TEXT
    $btnOpen.Font = $smallFont
    $btnOpen.Add_Click({
        if ($script:UfInstance -and (Test-Path -LiteralPath $script:UfInstance)) {
            Start-Process -FilePath explorer.exe -ArgumentList ('"' + $script:UfInstance + '"')
        }
    })
    $dlg.Controls.Add($btnOpen)
    $script:UfOpenBtn = $btnOpen

    $sep = New-Object System.Windows.Forms.Label
    $sep.AutoSize = $false
    $sep.Location = New-Object System.Drawing.Point(20, 120)
    $sep.Size = New-Object System.Drawing.Size(780, 2)
    $sep.BackColor = $C_PANEL
    $dlg.Controls.Add($sep)

    # 下方四個選項：各自顯示「該版本」下的完整目標路徑與統計
    $script:UfOptBtn = @{}
    $script:UfOptPath = @{}
    $script:UfOptStat = @{}
    $specs = @(
        @{ K='output'; C='Uf.ClearOutput'; D='Uf.DescOutput' }
        @{ K='custom'; C='Uf.ClearCustom'; D='Uf.DescCustom' }
        @{ K='temp';   C='Uf.ClearTemp';   D='Uf.DescTemp'   }
        @{ K='all';    C='Uf.ClearAll';    D='Uf.DescAll'    }
    )
    $y = 132
    foreach ($sp in $specs) {
        $hot = ($sp.K -eq 'custom' -or $sp.K -eq 'all')
        $b = New-Object System.Windows.Forms.Button
        $b.Text = L $sp.C
        $b.Location = New-Object System.Drawing.Point(20, $y)
        $b.Size = New-Object System.Drawing.Size(200, 32)
        $b.FlatStyle = 'Flat'
        $b.BackColor = $C_FIELD
        $b.ForeColor = $(if ($hot) { $C_WARN } else { $C_TEXT })
        $b.Font = $smallFont
        $b.Tag = $sp.K
        $b.Add_Click({
            $kind = [string]$args[0].Tag
            if ($script:UfInstance) { Invoke-UfAction -Kind $kind -Instance $script:UfInstance }
        })
        $dlg.Controls.Add($b)
        $script:UfOptBtn[$sp.K] = $b
        $script:UfTip.SetToolTip($b, (L $sp.D))

        $p = New-Object System.Windows.Forms.Label
        $p.AutoSize = $false
        $p.Location = New-Object System.Drawing.Point(232, ($y + 1))
        $p.Size = New-Object System.Drawing.Size(560, 18)
        $p.Font = $smallFont
        $p.ForeColor = $C_TEXT
        $p.AutoEllipsis = $true
        $dlg.Controls.Add($p)
        $script:UfOptPath[$sp.K] = $p

        $st = New-Object System.Windows.Forms.Label
        $st.AutoSize = $false
        $st.Location = New-Object System.Drawing.Point(232, ($y + 22))
        $st.Size = New-Object System.Drawing.Size(560, 18)
        $st.Font = $smallFont
        $st.ForeColor = $(if ($hot) { $C_WARN } else { $C_MUTED })
        $dlg.Controls.Add($st)
        $script:UfOptStat[$sp.K] = $st
        $y += 68
    }

    $note = New-Object System.Windows.Forms.Label
    $note.Text = L 'Dlg.UfNote'
    $note.AutoSize = $false
    $note.Location = New-Object System.Drawing.Point(20, 412)
    $note.Size = New-Object System.Drawing.Size(780, 38)
    $note.Font = $smallFont
    $note.ForeColor = $C_WARN
    $dlg.Controls.Add($note)

    $lblSt = New-Object System.Windows.Forms.Label
    $lblSt.Text = ''
    $lblSt.AutoSize = $false
    $lblSt.Location = New-Object System.Drawing.Point(20, 456)
    $lblSt.Size = New-Object System.Drawing.Size(620, 22)
    $lblSt.ForeColor = $C_TEXT
    $dlg.Controls.Add($lblSt)
    $script:UfStatus = $lblSt

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = L 'Dlg.UfClose'
    $btnClose.Location = New-Object System.Drawing.Point(688, 480)
    $btnClose.Size = New-Object System.Drawing.Size(112, 32)
    $btnClose.FlatStyle = 'Flat'
    $btnClose.BackColor = $C_FIELD
    $btnClose.ForeColor = $C_TEXT
    $btnClose.Add_Click({ $script:UfDlg.Close() })
    $dlg.Controls.Add($btnClose)
    $script:UfCloseBtn = $btnClose

    $cboVer.Add_SelectedIndexChanged({ Update-UfSelection })
    Refresh-UfInstances $Root ''

    if ($NoShow) { return $dlg }
    [void]$dlg.ShowDialog()
    $dlg.Dispose()
    try { $script:UfTip.Dispose() } catch { }
    $script:UfTip = $null
    $script:UfCbo = $null; $script:UfLoc = $null; $script:UfDlg = $null
    $script:UfVerLabel = $null; $script:UfHint = $null
    $script:UfOpenBtn = $null; $script:UfCloseBtn = $null
    $script:UfStatus = $null; $script:UfInstance = ''; $script:UfPaths = @()
    $script:UfOptBtn = @{}; $script:UfOptPath = @{}; $script:UfOptStat = @{}
    return $null
}

$btnUserFolder.Add_Click({ Show-UserFolderCleaner })

$script:DeleteAction = {
    $seen = @{}
    $uniq = New-Object System.Collections.ArrayList
    foreach ($it in $lvResults.CheckedItems) {
        $t = [string]$it.Tag
        if (-not $seen.ContainsKey($t)) { $seen[$t] = $true; [void]$uniq.Add($t) }
    }
    if ($uniq.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show((L 'Msg.NoSelection'), (L 'Msg.InfoTitle'), 'OK', 'Information') | Out-Null
        return
    }

    # 依長度排序後只查祖先鏈是否已入選：O(n x 深度)，取代原本 O(n²) 的兩層迴圈
    $targets = $uniq.ToArray() | Sort-Object { $_.Length }
    $kept = New-Object System.Collections.ArrayList
    $keptSet = @{}
    foreach ($t in $targets) {
        $covered = $false
        $parent = [System.IO.Path]::GetDirectoryName($t)
        while (-not [string]::IsNullOrEmpty($parent)) {
            if ($keptSet.ContainsKey($parent)) { $covered = $true; break }
            $parent = [System.IO.Path]::GetDirectoryName($parent)
        }
        if (-not $covered) { [void]$kept.Add($t); $keptSet[$t] = $true }
    }
    $skippedNested = $targets.Count - $kept.Count

    $protected = @()
    $final = @()
    foreach ($t in $kept) {
        if (Test-ProtectedPath -Path $t) { $protected += $t } else { $final += $t }
    }

    if ($protected.Count -gt 0) {
        Write-Log 'WARN' ((L 'Log.Protected') -f ($protected -join ' | '))
        [System.Windows.Forms.MessageBox]::Show(
            ((L 'Msg.ProtectedBody') + ($protected -join "`r`n")),
            (L 'Msg.ProtectedTitle'), 'OK', 'Warning') | Out-Null
    }
    if ($final.Count -eq 0) {
        $lblStatus.Text = L 'Status.NoDelete'
        return
    }

    $permanent = $chkPermanent.Checked
    $ans = Show-DeleteConfirm -Paths $final -Permanent $permanent
    if ($ans -ne 'ok') {
        Write-Log 'INFO' ((L 'Log.DeleteCancel') -f $final.Count)
        $lblStatus.Text = L 'Status.Cancelled'
        return
    }

    $mode = if ($permanent) { L 'Mode.Permanent' } else { L 'Mode.Recycle' }
    Write-Log 'INFO' ((L 'Log.DeleteStart') -f $mode, $final.Count, $skippedNested)

    $okCount = 0
    $failList = New-Object System.Collections.ArrayList
    $idx = 0
    $btnDelete.Enabled = $false
    $script:Deleting = $true
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    # 進度條：刪除中用紅色（危險），完成時轉綠色
    $script:ProgColor = $C_DANGER
    $script:ProgPct = 0
    $pnlProg.Height = 24
    $pnlProg.Invalidate()

    $deletedSet = @{}
    $script:LogBatching = $true
    $uiClock = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        foreach ($p in $final) {
            if ($script:AbortAll) { break }      # 檢查點 1：進入下一筆之前
            $idx++
            # UI 節流：每 50 筆或每 120ms 才更新狀態／進度／DoEvents
            if (($idx % 50) -eq 0 -or $uiClock.ElapsedMilliseconds -ge 120) {
                $lblStatus.Text = (L 'Status.Deleting') -f $idx, $final.Count, $p
                Set-DeleteProgress -Done $idx -Total $final.Count
                [System.Windows.Forms.Application]::DoEvents()
                $uiClock.Restart()
                if ($script:AbortAll) { break }  # 檢查點 2：DoEvents 可能剛處理了「立即停止」
            }
            $r = Remove-OneItem -Path $p -Permanent $permanent
            if ($r.Ok) {
                $okCount++
                $deletedSet[$p] = $true
                Write-Log 'OK' ((L 'Log.DeleteOk') -f $mode, $p)
            } else {
                Write-Log 'FAIL' ((L 'Log.DeleteFail') -f $mode, $p, $r.Reason)
                [void]$failList.Add(("{0}`r`n    → {1}" -f $p, $r.Reason))
            }
        }
    } finally {
        $script:LogBatching = $false
        Flush-Log
    }

    # 一次性移除成功刪除的列：原寫法每刪一筆就掃一遍 CheckedItems（O(n²)）
    if ($deletedSet.Count -gt 0) {
        $lvResults.BeginUpdate()
        try {
            for ($i = $lvResults.Items.Count - 1; $i -ge 0; $i--) {
                if ($deletedSet.ContainsKey([string]$lvResults.Items[$i].Tag)) { $lvResults.Items.RemoveAt($i) }
            }
        } finally { $lvResults.EndUpdate() }
    }
    Set-DeleteProgress -Done $idx -Total $final.Count

    $script:Deleting = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    $btnDelete.Enabled = $true

    # 進度條收尾：完成→綠色 100%；中斷→紅色停在實際完成比例
    if ($script:AbortAll) {
        Set-DeleteProgress -Done $idx -Total $final.Count
    } else {
        $script:ProgColor = $C_OK
        $script:ProgPct = 100
        $pnlProg.Invalidate()
    }

    if ($script:AbortAll) {
        # 立即停止：略過結果彈窗，記錄中斷位置後直接結束程式
        Write-Log 'WARN' ((L 'Log.DeleteAbort') -f ($okCount + $failList.Count), $final.Count, $okCount, $failList.Count, ($final.Count - $okCount - $failList.Count))
        try { $form.Close() } catch { }
        return
    }

    $summary = (L 'Status.Summary') -f $okCount, $failList.Count, $mode
    $lblStatus.Text = $summary
    Write-Log 'INFO' ((L 'Log.DeleteEnd') -f $summary)

    if (-not $script:PerfMode) {
        if ($failList.Count -gt 0) {
            $text = ((L 'Msg.FailListBody') -f $summary) + ($failList -join "`r`n`r`n")
            [System.Windows.Forms.MessageBox]::Show($text, (L 'Msg.PartialTitle'), 'OK', 'Warning') | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show(((L 'Msg.ResultBody') -f $summary), (L 'Msg.ResultTitle'), 'OK', 'Information') | Out-Null
        }
    }
    $pnlProg.Height = 0     # 結果已看過，收起進度條
}
$btnDelete.Add_Click($script:DeleteAction)

# ---------------------------------------------------------------- 關閉
# Ctrl+Q = 立即停止（KeyPreview 讓快捷鍵在任何控制項聚焦時都有效）
$form.KeyPreview = $true
$form.Add_KeyDown({
    if ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::Q) {
        $_.SuppressKeyPress = $true
        Invoke-AbortAll
    }
})

$form.Add_FormClosing({
    $script:AbortAll = $true   # 若關閉發生在刪除迴圈的 DoEvents 內，讓迴圈在檢查點停下
    try {
        if ($script:State) { $script:State.Cancel = $true }
        try { $script:Timer.Stop() } catch { }
        if ($script:PS) { try { $script:PS.Stop() } catch { }; try { $script:PS.Dispose() } catch { } }
    } catch { }
    Write-Log 'INFO' (L 'Log.Close')
    Flush-Log
})

Write-Log 'INFO' (L 'Log.Start')
$lblStatus.Text = L 'Status.ReadyFull'

if ($SelfTest) {
    [void]$form.CreateControl()
    $line = 'SELFTEST controls={0} top={1} left={2} bottom={3} list={4} status={5} drives={6} recycle={7}' -f `
        $form.Controls.Count, $pnlTop.Bounds, $pnlLeft.Bounds, $pnlBottom.Bounds, $lvResults.Bounds, $status.Bounds, $clbDrives.Items.Count, $script:RecycleOk
    $line2 = 'SELFTEST2 topW={0} ext={1} extRight={2} hintRight={3} zwScope={4} ribbon={5} searchBtn={6} uiCbo={7} langCbo={8}' -f `
        $pnlTop.Width, $cboExt.Bounds, ($cboExt.Left + $cboExt.Width), ($lblZw3dHint.Left + $lblZw3dHint.Width), $cboZw3dScope.Bounds, $pnlRibbon.Bounds, $btnSearch.Bounds, $cboUi.Bounds, $cboLang.Bounds
    $line3 = 'SELFTEST3 uiLang={0} scanLang={1} keys={2} words={3} presets={4} title={5} uiLabel=[{6}] scanLabel=[{7}]' -f `
        $script:UiLang, $(if ($script:ScanLangId) { $script:ScanLangId } else { 'All' }), $script:Str.Count, @(Get-LangWords $script:ScanLangId).Count, $cboPreset.Items.Count, $form.Text, $lblUi.Text, $lblLang.Text
    Write-Log 'INFO' $line
    Write-Log 'INFO' $line2
    Write-Log 'INFO' $line3
    $line
    $line2
    $line3

    # 語言版本切換的搜尋規則自檢：同一組目錄，不同語言版本應得到不同命中數
    $testRoot = Join-Path $script:BaseDir ('FM_LANG_ZW3D_' + [guid]::NewGuid().ToString('N'))
    [void][System.IO.Directory]::CreateDirectory($testRoot)
    try {
        foreach ($d in @('備份', 'バックアップ', 'Backup', 'Sicherung', '一般')) {
            [void][System.IO.Directory]::CreateDirectory((Join-Path $testRoot $d))
        }
        [void][System.IO.File]::WriteAllText((Join-Path $testRoot 'A.Z3PRT.1.z3bak'), 'x')
        $expect = @{ '' = 4; 'zh_TW' = 1; 'zh_CN' = 0; 'ja_JP' = 1; 'de_DE' = 1; 'en_US' = 1 }
        $got = @()
        foreach ($id in @('', 'zh_TW', 'zh_CN', 'ja_JP', 'de_DE', 'en_US')) {
            $st = [hashtable]::Synchronized(@{ Done=$false; Cancel=$false; Scanning=''; Scanned=0; Errors=0; Results=$null; Truncated=$false })
            & $searchBlock @($testRoot) '' 0 5000 3 2 '' $st @(Get-LangWords $id) 'File' 'Folder'
            $n = if ($st.Results) { $st.Results.Count } else { 0 }
            $got += ('{0}={1}' -f $(if ($id) { $id } else { 'All' }), $n)
            if ($n -ne $expect[$id]) { throw ('SELFTEST lang {0}: expected {1}, got {2}' -f $id, $expect[$id], $n) }
        }
        # TargetIndex=1（檔案）與 Type 穩定值，驗證下拉文字解耦後邏輯未壞
        $stF = [hashtable]::Synchronized(@{ Done=$false; Cancel=$false; Scanning=''; Scanned=0; Errors=0; Results=$null; Truncated=$false })
        & $searchBlock @($testRoot) '' 1 5000 3 2 '' $stF @() 'File' 'Folder'
        if ($stF.Results.Count -ne 1 -or [string]$stF.Results[0].Type -ne 'File') {
            throw 'SELFTEST file mode / Type decoupling failed'
        }
        $got += 'Files=1'
        $line4 = 'SELFTEST4 langScan={0}' -f ($got -join ',')
        Write-Log 'INFO' $line4
        $line4
    } finally {
        try { [System.IO.Directory]::Delete($testRoot, $true) } catch { }
    }

    # 執行期切換自檢：兩個下拉各自獨立，互不干擾
    $origUi = $cboUi.SelectedIndex
    $origScan = $cboLang.SelectedIndex
    $scanBefore = $script:ScanLangId
    $uiChecks = @{ 2 = 'ZW3D Files'; 13 = '文件管理工具'; 14 = '檔案管理工具' }
    $switchLog = @()
    foreach ($i in @(2, 13, 14)) {
        $cboUi.SelectedIndex = $i
        $id = $script:Langs[$i].Id
        if ($form.Text.IndexOf($uiChecks[$i]) -lt 0 -or $script:UiLang -ne $id) {
            throw ('SELFTEST ui switch failed at index {0}' -f $i)
        }
        if ($script:ScanLangId -ne $scanBefore) { throw 'SELFTEST ui switch changed scan language' }
        # 版面守門：第三列的提示、兩個標籤與兩個下拉不可互相重疊，也不可超出頂部工具列
        if ($lblHint.Right -gt $lblUi.Left -or $lblUi.Right -gt $cboUi.Left -or
            $cboUi.Right -gt $lblLang.Left -or $lblLang.Right -gt $cboLang.Left -or
            $cboLang.Right -gt $pnlTop.Width) {
            throw ('SELFTEST layout overlap in {0}: hint={1} uiLbl={2} uiCbo={3} scanLbl={4} scanCbo={5} topW={6}' -f `
                $id, $lblHint.Right, $lblUi.Right, $cboUi.Right, $lblLang.Right, $cboLang.Right, $pnlTop.Width)
        }
        $switchLog += ('ui={0}' -f $id)
    }
    # 掃描語言版本切換：只動預設詞，不動介面語言
    $uiBefore = $script:UiLang
    $cboLang.SelectedIndex = 14                  # zh_CN
    if ($cboPreset.Items -notcontains '备份') { throw 'SELFTEST localized presets failed' }
    if ($script:UiLang -ne $uiBefore -or $script:ScanLangId -ne 'zh_CN') { throw 'SELFTEST scan switch coupling failed' }
    $switchLog += ('scan={0}:presets={1}' -f $script:ScanLangId, $cboPreset.Items.Count)
    $cboLang.SelectedIndex = $origScan
    # 多國語言覆蓋率：每種介面語言都必須有全部畫面／視窗標題文案（日誌與語言中立詞條除外）
    $neutral = @('Export.FilterCsv', 'Export.FilterJson', 'Ext.Unlimited')
    $missing = @()
    foreach ($lg in $script:Langs) {
        foreach ($k in $script:Str.Keys) {
            if ($k.StartsWith('Log.') -or $neutral -contains $k) { continue }
            if (-not $script:Str[$k][$lg.Id]) { $missing += ('{0}:{1}' -f $lg.Id, $k) }
        }
    }
    if ($missing.Count -gt 0) { throw ('SELFTEST missing translations: ' + (($missing | Select-Object -First 20) -join ', ')) }
    # 逐一切換全部介面語言：視窗標題必須跟著換，版面不可重疊
    for ($i = 0; $i -lt $script:Langs.Count; $i++) {
        $cboUi.SelectedIndex = $i
        if ($form.Text -ne (L 'AppTitle') -or $lblStatus.Text -eq '' -or $cboLang.Items[0] -ne (L 'Lang.All')) {
            throw ('SELFTEST title/lang switch failed for {0}' -f $script:Langs[$i].Id)
        }
        if ($lblHint.Right -gt $lblUi.Left) { throw ('SELFTEST hint overlaps language label in {0}' -f $script:Langs[$i].Id) }
    }
    $switchLog += ('all={0}' -f $script:Langs.Count)
    $cboUi.SelectedIndex = $origUi
    $line5 = 'SELFTEST5 langSwitch={0}' -f ($switchLog -join ',')
    Write-Log 'INFO' $line5
    $line5

    # 使用者目錄清理：合成假實例目錄，驗證 4 個動作的目標與結果，绝不动真实目录
    $ufRoot = Join-Path $script:BaseDir ('FM_UF_' + [guid]::NewGuid().ToString('N'))
    $ufInst = Join-Path $ufRoot 'ZW3D Test 2027'
    try {
        foreach ($sub in @('custom', 'output', 'temp')) {
            $sd = Join-Path $ufInst $sub
            [void][System.IO.Directory]::CreateDirectory($sd)
            [void][System.IO.File]::WriteAllText((Join-Path $sd 'a.txt'), 'x')
        }
        $found = @(Get-UserFolderInstances $ufRoot)
        if ($found.Count -ne 1 -or $found[0] -ne $ufInst) { throw 'SELFTEST7 instance detection failed' }
        foreach ($k in @('output', 'custom', 'temp')) {
            $expect = Join-Path $ufInst $k
            if ((Get-UserCleanTarget $ufInst $k) -ne $expect) { throw ('SELFTEST7 target {0} failed' -f $k) }
        }
        if ((Get-UserCleanTarget $ufInst 'all') -ne $ufInst) { throw 'SELFTEST7 target all failed' }

        $ufLog = @()
        $done = @{}
        foreach ($k in @('output', 'custom', 'temp')) {
            $r = Invoke-UserClean -Instance $ufInst -Kind $k -Permanent $true
            if (-not $r.Ok -or (Test-Path -LiteralPath (Join-Path $ufInst $k))) {
                throw ('SELFTEST7 clean {0} failed' -f $k)
            }
            $done[$k] = $true
            foreach ($other in @('custom', 'output', 'temp')) {
                if ($done.ContainsKey($other)) { continue }
                if (-not (Test-Path -LiteralPath (Join-Path $ufInst $other))) {
                    throw ('SELFTEST7 clean {0} removed sibling {1}' -f $k, $other)
                }
            }
            $ufLog += ($k + '=ok')
        }
        $rAll = Invoke-UserClean -Instance $ufInst -Kind 'all' -Permanent $true
        if (-not $rAll.Ok -or (Test-Path -LiteralPath $ufInst)) { throw 'SELFTEST7 clean all failed' }
        $ufLog += 'all=ok'
        $miss = Invoke-UserClean -Instance $ufInst -Kind 'temp' -Permanent $true
        if (-not $miss.Ok -or -not $miss.Skipped) { throw 'SELFTEST7 missing target handling failed' }
        $ufLog += 'missing=skip'
        $line7 = 'SELFTEST7 userFolder={0}' -f ($ufLog -join ',')
        Write-Log 'INFO' $line7
        $line7
    } finally {
        try { [System.IO.Directory]::Delete($ufRoot, $true) } catch { }
    }

    # 新按鈕必須在工具列內，且四個動作的文案在目前語言下都取得到
    $ufBtnOk = ($btnUserFolder -and $btnUserFolder.Width -gt 0 -and
                ($btnUserFolder.Left + $btnUserFolder.Width) -le $pnlRibbon.Width)
    $ufKeys = @('Uf.ClearOutput', 'Uf.ClearCustom', 'Uf.ClearTemp', 'Uf.ClearAll',
                'Btn.UserFolder', 'Dlg.UfTitle', 'Uf.Confirm',
                'Dlg.UfVersion', 'Dlg.UfLocation', 'Dlg.UfHint', 'Dlg.UfOpen',
                'Dlg.UfClose', 'Uf.StatLine', 'Uf.NoTarget')
    foreach ($k in $ufKeys) { if ((L $k) -eq $k) { throw ('SELFTEST8 missing key {0}' -f $k) } }
    if (-not $ufBtnOk) { throw 'SELFTEST8 user-folder button out of ribbon' }
    $line8 = 'SELFTEST8 ufButton=ok ufKeys={0}' -f $ufKeys.Count
    Write-Log 'INFO' $line8
    $line8

    # 版本下拉：自動帶出本機版本目錄，切換即更新位置與四個選項的完整路徑
    $ufRoot9 = Join-Path $script:BaseDir ('FM_UFV_' + [guid]::NewGuid().ToString('N'))
    try {
        $v1 = Join-Path $ufRoot9 'ZW3D 2026'
        $v2 = Join-Path $ufRoot9 'ZW3D WuKong 2027'
        foreach ($v in @($v1, $v2)) {
            foreach ($sub in @('custom', 'output', 'temp')) {
                $sd = Join-Path $v $sub
                [void][System.IO.Directory]::CreateDirectory($sd)
                [void][System.IO.File]::WriteAllText((Join-Path $sd 'a.txt'), 'xxxx')
            }
        }
        $vit = @(Get-UfVersionItems $ufRoot9)
        if ($vit.Count -ne 2) { throw 'SELFTEST9 version detection failed' }
        foreach ($x in $vit) {
            if ($x.Name -ne (Split-Path -Leaf $x.Path) -or $x.Name -match '\\') {
                throw 'SELFTEST9 version item is not a plain version name'
            }
        }
        $d9 = Show-UserFolderCleaner -Root $ufRoot9 -NoShow
        if (-not $d9) { throw 'SELFTEST9 dialog creation failed' }
        try {
            if ($script:UfCbo.Items.Count -ne 2) { throw 'SELFTEST9 combo count failed' }
            foreach ($it in $script:UfCbo.Items) {
                if ([string]$it -match '\\') { throw 'SELFTEST9 combo shows a path instead of a version' }
            }
            if (-not $script:UfLoc.ReadOnly) { throw 'SELFTEST9 location is editable' }
            $p0 = [string]$script:UfPaths[0]
            if ($script:UfLoc.Text -ne $p0) { throw 'SELFTEST9 initial location failed' }
            foreach ($k in @('output', 'custom', 'temp', 'all')) {
                if (-not $script:UfOptPath[$k].Text.StartsWith($p0, [System.StringComparison]::OrdinalIgnoreCase)) {
                    throw ('SELFTEST9 option path {0} failed' -f $k)
                }
            }
            $script:UfCbo.SelectedIndex = 1
            $p1 = [string]$script:UfPaths[1]
            if ($script:UfLoc.Text -ne $p1) { throw 'SELFTEST9 location did not follow version switch' }
            if ($p1 -eq $p0) { throw 'SELFTEST9 two versions share a path' }
            foreach ($k in @('output', 'custom', 'temp', 'all')) {
                if (-not $script:UfOptPath[$k].Text.StartsWith($p1, [System.StringComparison]::OrdinalIgnoreCase)) {
                    throw ('SELFTEST9 option path {0} did not follow switch' -f $k)
                }
            }
            if ($script:UfOptStat['all'].Text -notmatch '\d') { throw 'SELFTEST9 stats missing' }
            if ($script:UfCbo.Top + $script:UfCbo.Height -ge $script:UfLoc.Top) { throw 'SELFTEST9 version/location overlap' }
            foreach ($k in @('output', 'custom', 'temp', 'all')) {
                $b9 = $script:UfOptBtn[$k]; $l9 = $script:UfOptPath[$k]
                if ($b9.Right -gt $l9.Left) { throw ('SELFTEST9 button/path overlap {0}' -f $k) }
                if ($l9.Right -gt $d9.ClientSize.Width) { throw ('SELFTEST9 path label outside form {0}' -f $k) }
            }
            if ($script:UfVerLabel.Right -gt $script:UfCbo.Left) { throw 'SELFTEST9 version label overlaps combo' }
            if ($script:UfOpenBtn.Right -gt $d9.ClientSize.Width) { throw 'SELFTEST9 open button outside form' }
            foreach ($c9 in $d9.Controls) {
                if ($c9.Right -gt $d9.ClientSize.Width -or $c9.Bottom -gt $d9.ClientSize.Height) {
                    throw ('SELFTEST9 control outside form: {0} @ {1},{2}' -f $c9.GetType().Name, $c9.Right, $c9.Bottom)
                }
            }
            $label9 = [string]$script:UfVerLabel.Text
            if (-not $label9 -or $label9 -eq 'Dlg.UfVersion') { throw 'SELFTEST9 version label missing' }
            $uf9Log = 'versions=2,loc=switch,paths=4,readonly=1,layout=ok,label=' + $label9
        } finally {
            $d9.Dispose()
            $script:UfCbo = $null; $script:UfLoc = $null; $script:UfDlg = $null
            $script:UfVerLabel = $null; $script:UfHint = $null; $script:UfTip = $null
            $script:UfStatus = $null; $script:UfInstance = ''; $script:UfPaths = @()
            $script:UfOptBtn = @{}; $script:UfOptPath = @{}; $script:UfOptStat = @{}
        }
        $line9 = 'SELFTEST9 ufVersion={0}' -f $uf9Log
        Write-Log 'INFO' $line9
        $line9
    } finally {
        try { [System.IO.Directory]::Delete($ufRoot9, $true) } catch { }
    }

    if ($Perf) {
        $perfRoot = Join-Path $script:BaseDir ('FM_PERF_' + [guid]::NewGuid().ToString('N'))
        [void][System.IO.Directory]::CreateDirectory($perfRoot)
        try {
            $sw = New-Object System.Diagnostics.Stopwatch
            $perfLog = @()

            # 1) 日誌：真實 Write-Log 1000 次
            $sw.Restart()
            for ($i = 0; $i -lt 1000; $i++) { Write-Log 'INFO' ('PERF log {0}' -f $i) }
            $sw.Stop(); $perfLog += ('log1000={0}' -f $sw.ElapsedMilliseconds)

            # 2) 結果填入 5000 筆：走真實 Complete-Search
            $fake = New-Object System.Collections.ArrayList
            for ($i = 0; $i -lt 5000; $i++) {
                [void]$fake.Add([pscustomobject]@{
                    Name = ('item_{0}.z3bak' -f $i)
                    Path = (Join-Path $perfRoot ('item_{0}.z3bak' -f $i))
                    Type = 'File'
                    Parent = $perfRoot
                })
            }
            $script:State = @{ Done=$true; Results=$fake; Scanned=5000; Errors=0;
                               Truncated=$false; Keyword=''; Target='Files' }
            $sw.Restart()
            Complete-Search
            $sw.Stop(); $perfLog += ('fill5000={0}' -f $sw.ElapsedMilliseconds)

            # 3) 勾選全部 5000 筆：走真實按鈕事件
            $sw.Restart()
            & $script:CheckAllAction
            $sw.Stop(); $perfLog += ('check5000={0}' -f $sw.ElapsedMilliseconds)

            # 4) 刪除 800 個暫存檔：走真實 btnDelete（含確認/進度/日誌/清單移除）
            $lvResults.BeginUpdate()
            $lvResults.Items.Clear()
            for ($i = 0; $i -lt 800; $i++) {
                $dp = Join-Path $perfRoot ('d_{0}.tmp' -f $i)
                [void][System.IO.File]::WriteAllText($dp, 'x')
                $di = New-Object System.Windows.Forms.ListViewItem(('d_{0}.tmp' -f $i))
                [void]$di.SubItems.Add('File')
                [void]$di.SubItems.Add($dp)
                [void]$di.SubItems.Add($perfRoot)
                $di.Tag = $dp
                $di.Checked = $true
                [void]$lvResults.Items.Add($di)
            }
            $lvResults.EndUpdate()
            for ($i = 0; $i -lt $lvResults.Items.Count; $i++) { $lvResults.Items[$i].Checked = $true }
            $chkPermanent.Checked = $true        # 固定永久刪除，排除回收筒 COM 的變異
            $sw.Restart()
            & $script:DeleteAction
            $sw.Stop(); $perfLog += ('delete800={0}' -f $sw.ElapsedMilliseconds)
            $left = @([System.IO.Directory]::EnumerateFiles($perfRoot, 'd_*.tmp')).Count
            $perfLog += ('delLeft={0}' -f $left)
            if ($left -ne 0) { throw ('PERF delete incomplete: {0} left' -f $left) }

            # 5) 搜尋：合成 20 目錄 x 50 檔案，走真實 searchBlock
            $scanRoot = Join-Path $perfRoot 'scan'
            [void][System.IO.Directory]::CreateDirectory($scanRoot)
            for ($d = 0; $d -lt 20; $d++) {
                $sd = Join-Path $scanRoot ('dir_{0}' -f $d)
                [void][System.IO.Directory]::CreateDirectory($sd)
                for ($f = 0; $f -lt 50; $f++) {
                    [void][System.IO.File]::WriteAllText((Join-Path $sd ('ZW3D_{0}.z3bak' -f $f)), 'x')
                }
            }
            $st = [hashtable]::Synchronized(@{ Done=$false; Cancel=$false; Scanning='';
                                               Scanned=0; Errors=0; Results=$null; Truncated=$false })
            $sw.Restart()
            & $searchBlock @($scanRoot) '' 0 5000 3 2 '' $st @(Get-LangWords $script:ScanLangId) 'File' 'Folder'
            $sw.Stop(); $perfLog += ('scan1000={0}' -f $sw.ElapsedMilliseconds)

            $line6 = 'PERF {0}' -f ($perfLog -join ' ')
            Write-Log 'INFO' $line6
            $line6
        } finally {
            try { [System.IO.Directory]::Delete($perfRoot, $true) } catch { }
        }
    }
    Flush-Log
    $form.Dispose()
} else {
    [void]$form.ShowDialog()
}