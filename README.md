# Explorer Merge

**Too many File Explorer windows? One click. One window.**

A small, free Windows 11 utility that reopens the tabs from separate File Explorer windows in one native tabbed window, then closes the verified source windows.

No installer. No background service. No third-party runtime. No network access. MIT licensed.

<img src="assets/ExplorerMerge-character.png" width="160" alt="Explorer Merge character holding a folder">

[紹介ページ / App page](https://www.aoe1928.com/apps/)

## 日本語

散らばったエクスプローラーを、クリック1回で1つのウィンドウにまとめます。

### 使い方

1. [Releases](https://github.com/aoe1928/explorer-merge/releases/latest) から ZIP をダウンロードして、好きな場所に展開します。
2. `ExplorerMerge.exe` を実行すると、現在の仮想デスクトップのエクスプローラーをタブにまとめます。
3. タスクバーから使う場合は、exe を右クリック →「その他のオプションを確認」→「タスクバーにピン留めする」。Windowsによっては最初の右クリックメニューに表示されます。

ピン留め後は exe の保存場所を変えないでください。管理者権限は不要です。常駐せず、処理が終わると終了します。通常のファイルやフォルダーの内容は変更しません。

### 動作と制限

- Windows 11 22H2（ビルド22621）以降の、標準タブ付きエクスプローラーが必要です。
- 現在の仮想デスクトップで手前にある、読み取り可能なウィンドウを結合先にします。最小化されたウィンドウも対象です。他の仮想デスクトップのウィンドウは対象外です。
- 元の各ウィンドウの全タブを開き直し、移し先の場所と元タブの状態を確認してから元ウィンドウを閉じます。確認できない場合は処理を止め、未処理の元ウィンドウを残します。
- 同じフォルダーを開いていたタブも、その数だけ残します。
- 結合中は結合先を一時的に最小化し、終了時に戻します。タブごとの画面切り替えを抑え、必要なタブだけを確認しながら複数フォルダーを読み込みます。
- **タブの「場所」を開き直す方式です。戻る・進むの履歴、選択中のファイル、スクロール位置、タブの並び順は引き継ぎません。**
- 実行中はエクスプローラーの操作を止め、完了まで数秒待ってください。大量のタブや遅い場所では長くかかります。
- 読み込み中・応答しないフォルダーや、一部の特殊フォルダーはまとめられない場合があります。エラー時には、新しく作ったタブが残る場合があります。
- Windowsの非公開タブ操作を使用しているため、Windows更新によって動かなくなる可能性があります。

### ダウンロードした exe を使わず、自分のPCで作る

ソースの ZIP を展開し、`Build.cmd` をダブルクリックしてください。Windows標準の .NET Framework コンパイラを使い、`dist\ExplorerMerge.exe` を作ります。Visual Studio、.NET SDK、追加ライブラリのインストールは不要です。

配布ファイルはコード署名していません。Windowsが警告する場合は、ソースを確認して手元でビルドすることもできます。

削除するときはタスクバーのピン留めを外し、展開したファイルを削除するだけです。設定、常駐プロセス、自動起動登録はありません。

## English

### Use

Download and extract the ZIP from [Releases](https://github.com/aoe1928/explorer-merge/releases/latest). Run `ExplorerMerge.exe` whenever your Explorer windows have multiplied. To make this a taskbar button, right-click the executable, choose **Show more options → Pin to taskbar** (menu placement varies). Keep the executable at the same path after pinning.

Requires Windows 11 22H2 or later with native Explorer tabs. Runs as a normal user and exits after merging. Minimized windows on the current virtual desktop are included; cloaked windows on other desktops are excluded. The first readable Explorer window in front-to-back order receives the tabs.

### What is preserved?

Folder locations and duplicate tabs are preserved. This **reopens locations**, rather than moving live tabs: navigation history, selected files, scroll positions and tab order are not preserved. Wait for the operation to finish before interacting with Explorer.

Every source window must have a complete readable tab snapshot. Each new tab's location is checked before the source frame is closed, and source tabs are checked again for changes. An uncertain result stops processing and leaves remaining source windows open. New tabs may remain after a failure. These checks reduce accidental loss of locations; they do not make concurrent interaction atomic.

Some shell namespaces, loading folders, unavailable network locations, Explorer extensions and future Windows builds may be unsupported. Tab creation uses an undocumented Explorer command. No `explorer.exe` processes are terminated, and no global keyboard input is injected.

### Performance (v0.2.0)

The destination is temporarily minimized during merging and restored on both success and failure. The program caches tab COM objects, navigates as soon as a new tab is registered instead of waiting for Home to finish loading, overlaps folder loads, and verifies only relevant locations. It no longer repeatedly reads every folder in every Explorer window while polling. Location checks before closing source windows remain in place.

This still reopens locations; it does not transplant live tabs or eliminate filesystem enumeration. Slow storage, network folders and Explorer extensions can still affect completion time.

On one Windows 11 25H2 desktop, six test windows containing 100 files per folder merged in **42.0 seconds with v0.1.0 versus 4.7–6.4 seconds with v0.2.0** (one baseline run, two optimized runs). These are observed desktop timings, not a guaranteed speedup on every PC. The benchmark verifies the complete set of locations and source-window closure.

### Build and test

```powershell
.\Build.ps1
.\Test.ps1
# Opens and closes only test-created Explorer windows; requires an interactive desktop:
.\Test.ps1 -Integration
# Six test windows with 100 files per folder; measures merge time only:
.\Test.ps1 -Benchmark
```

Build uses the C# compiler and .NET Framework included with 64-bit Windows. It requires no NuGet packages or SDK downloads. The executable is unsigned. CI builds the executable and runs the safety-policy tests; native Explorer integration must be tested on a real desktop.

The character icon is based on original artwork by aoe1928, adapted with image generation. `tools/PackIcon.ps1` packages the selected PNG into the multi-resolution ICO used by the build; the original geometric icon is retained in `assets/ExplorerMerge.ico`.

Diagnostic options (intended for development):

```text
ExplorerMerge.exe --inspect snapshot.xml
ExplorerMerge.exe --windows 123456,234567 --report result.xml
```

`--inspect` explicitly exports local folder paths and window handles. Do not upload this file without reviewing it. Normal execution writes no diagnostic files and sends nothing over the network. `--windows` limits merging to the specified current, eligible window handles, with the first as the destination. Exit codes: 0 success/no-op, 1 fatal error, 2 partial/skipped operation.

### Contributing

Please include your Windows build and a reproduction using non-private test folders. Do not attach personal folder snapshots. Run the policy tests and, for tab-management changes, the interactive integration tests.

### Implementation references

- [Microsoft: IShellBrowser](https://learn.microsoft.com/windows/win32/api/shobjidl_core/nn-shobjidl_core-ishellbrowser)
- [Microsoft: SHGetIDListFromObject](https://learn.microsoft.com/windows/win32/api/shlobj_core/nf-shlobj_core-shgetidlistfromobject)
- [ExplorerTabUtility](https://github.com/w4po/ExplorerTabUtility): reference for the shell-browser service identifier and the Explorer new-tab command. Explorer Merge has an independent, one-shot implementation.

## License

[MIT](LICENSE). Free to use, modify and redistribute.
