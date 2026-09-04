# ============================================
# スリジエ スタッツ自動集計ツール（第1弾）
# VolleyStationの .dvw ファイルから、
# アタックの「打数・得点・ミス・決定率・効果率」を
# 選手ごと・セットごとに自動計算します
# ============================================

import os
import re
import json

# ---- ここだけ書き換えれば別の試合・別のPCでも動きます ----
FILE_PATH = r"C:\Users\kouse\OneDrive\添付ファイル\東京スリジエ\2026-07-19 TSJ-KOU.dvw"

# 対戦相手の名前（.dvwファイルの中では文字化けして読めないので手入力）
OPPONENT_NAME = "江戸川大学"

# 'women'（女子）か 'men'（男子）かをここで切り替える。
# 男子の試合を読み込むときは 'men' に変えるだけで、名簿(ROSTER)・
# 出力フォルダ・ページのタイトルが自動的に男子用に切り替わる
TEAM = 'women'

# ---- 映像連携（プレー動画へのリンク）用の設定。使わない試合は空のままでOK ----
# YouTube（限定公開）などにアップロードした、その試合の動画のURL。
# 空のままにしておくと、ダッシュボード側で映像リンクの表示自体が出なくなる。
VIDEO_URL = ''
# 「動画の何秒目が、.dvwの中の時刻（壁時計）の何時何分何秒に対応するか」の差分（秒）。
# 動画を実際に見て、分かっているプレーの動画上の秒数から、
# 動画秒数 - その行の壁時計の秒数（真夜中から数えて） で計算する。
# 分からない場合はNoneのままにしておくと、映像リンクは計算されない。
VIDEO_OFFSET_SECONDS = None

# ※出力ファイル名は「試合日_対戦相手.html」の形で自動的に決まるので、
# ここを手で書き換える必要はもうありません（build_output_filename関数を参照）
# 保存先も women/ または men/ フォルダに自動で分かれます

# 背番号 → 名前 の対応表。チームごとに分けてあるので、TEAMの設定で自動的に選ばれる
ROSTER_BY_TEAM = {
    'women': {
        1: '本村 嘉菜',
        2: '北川 茉奈',
        3: '島崎 葵',
        4: '菊田 美優',
        5: '萩谷 歩未',
        6: '大石 麻美',
        7: '坂本 朱乃',
        8: '早野 有美',
        9: '佐藤 まひろ',
        10: '関口 希望',
        11: '安藤 美果',
        12: '小林 琉夏',
    },
    # ↓ .dvwファイルの中のローマ字名簿と、こうせいさんが送ってくれた
    # 「スリジエ男子データ.xlsx」の漢字名簿を、名前を手がかりに突き合わせたものです。
    # 一部（4,7,8,13,15,18,22,27番）はExcel側に見当たらなかったのでローマ字のまま、
    # 17・25・26番は苗字の一致だけで判断した推測なので、確認をお願いします
    'men': {
        1: '大楽 祥生',
        2: '浅岡 遥太',  # リベロ
        3: '中島 博雅',
        4: 'NAKANO SOUICHIROU',  # ← Excelに見当たらず。漢字が分かれば教えてください
        5: '中村 耕平',
        6: '岡田 大雅',
        7: 'YAMAZAKI SATOSHI',  # ← Excelに見当たらず
        8: 'YUUSUKE AOKI',  # ← Excelに見当たらず
        9: '佐藤 匠',
        10: '勝俣 樹生',
        11: '奥田 晃',
        12: '渡邊 颯',
        13: 'TOMOYA FUJITA',  # ← Excelに見当たらず（リベロ）
        14: '宮下 ジェイラン',
        15: 'SATO KOUSUKE',  # ← Excelに見当たらず
        17: '久保 昴也',  # ← 苗字のみ一致（推測）
        18: 'OHSHITA TADASHI',  # ← Excelに見当たらず
        22: 'SHIRAI HIROTO',  # ← Excelに見当たらず
        23: '南 賢清',
        25: '久本 快',  # ← 推測
        26: '奥田 祥太',  # ← 推測
        27: 'TAKAOKA KAKERU',  # ← Excelに見当たらず
        31: '斎藤 直樹',
    },
}

TEAM_LABELS = {'women': '東京スリジエ（女子）', 'men': '東京スリジエ（男子）'}
PAGE_TITLES = {'women': '東京スリジエ女子試合データ', 'men': '東京スリジエ男子試合データ'}

ROSTER = ROSTER_BY_TEAM[TEAM]
TEAM_LABEL = TEAM_LABELS[TEAM]


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__PAGE_TITLE__</title>
<style>
  .viz-root {
    color-scheme: light;
    --surface-1:      #fcfcfb;
    --page-plane:     #f9f9f7;
    --text-primary:   #0b0b0b;
    --text-secondary: #52514e;
    --text-muted:     #898781;
    --grid:           #e1e0d9;
    --baseline:       #c3c2b7;
    --border:         rgba(11,11,11,0.10);
    --series-blue:    #2a78d6;
    --series-orange:  #eb6834;
    --series-aqua:    #1baf7a;
    --series-yellow:  #eda100;
  }
  @media (prefers-color-scheme: dark) {
    .viz-root {
      color-scheme: dark;
      --surface-1:      #1a1a19;
      --page-plane:     #0d0d0d;
      --text-primary:   #ffffff;
      --text-secondary: #c3c2b7;
      --text-muted:     #898781;
      --grid:           #2c2c2a;
      --baseline:       #383835;
      --border:         rgba(255,255,255,0.10);
      --series-blue:    #3987e5;
      --series-orange:  #d95926;
      --series-aqua:    #199e70;
      --series-yellow:  #c98500;
    }
  }

  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
    background: var(--page-plane);
    color: var(--text-primary);
  }
  .wrap { max-width: 1180px; margin: 0 auto; padding: 32px 20px 80px; }
  header { margin-bottom: 28px; }
  header h1 { font-size: 30px; margin: 0 0 4px; }
  header p { margin: 0; color: var(--text-secondary); font-size: 14px; }

  .sidebar-toggle-btn {
    font-family: inherit; font-size: 16px; font-weight: 600; cursor: pointer;
    width: 40px; height: 40px; border-radius: 10px;
    border: 1px solid var(--border); background: var(--surface-1); color: var(--text-primary);
    display: flex; align-items: center; justify-content: center; margin-bottom: 14px;
  }
  .sidebar-toggle-btn:hover { background: var(--grid); }

  .app-layout { display: flex; align-items: flex-start; gap: 24px; }
  .sidebar-column { width: 300px; flex-shrink: 0; position: sticky; top: 20px; }
  .sidebar.collapsed { display: none; }
  .sidebar .card { margin-bottom: 18px; }
  .sidebar .card:last-child { margin-bottom: 0; }
  .sidebar-label { font-size: 12.5px; color: var(--text-muted); margin-bottom: 8px; }
  .main-content { flex: 1; min-width: 0; }
  @media (max-width: 860px) {
    .app-layout { flex-direction: column; }
    .sidebar-column { width: 100%; position: static; }
  }

  .player-toggle-btn {
    width: 100%; font-family: inherit; font-size: 13px; font-weight: 600; cursor: pointer;
    padding: 9px 12px; border-radius: 8px; margin-top: 8px;
    border: 1px solid var(--border); background: var(--page-plane); color: var(--text-primary);
    display: flex; align-items: center; justify-content: space-between; gap: 8px;
  }
  .player-toggle-btn:hover { background: var(--grid); }
  .player-toggle-btn .arrow { color: var(--text-muted); font-size: 11px; }
  .player-checklist {
    display: flex; flex-direction: column; gap: 2px; margin-top: 8px;
    max-height: 320px; overflow-y: auto;
  }
  .player-check-item {
    display: flex; align-items: center; gap: 9px; font-size: 13.5px;
    padding: 7px 6px; border-radius: 6px; cursor: pointer; color: var(--text-primary);
  }
  .player-check-item:hover { background: var(--grid); }
  .player-check-item input[type="checkbox"] { width: 16px; height: 16px; accent-color: var(--series-blue); cursor: pointer; }
  .match-row { display: flex; align-items: center; gap: 2px; }
  .match-row .player-check-item { flex: 1; min-width: 0; }
  .match-row .player-check-item span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .match-summary-icon-btn {
    flex-shrink: 0; border: none; background: none; cursor: pointer; font-size: 15px;
    padding: 6px 8px; border-radius: 6px; color: var(--text-muted); line-height: 1;
  }
  .match-summary-icon-btn:hover { background: var(--grid); color: var(--series-blue); }
  .match-month-header {
    display: flex; align-items: center; justify-content: space-between; gap: 8px;
    padding: 8px 6px 4px; font-size: 12px; font-weight: 700; color: var(--text-muted);
  }
  .match-month-header:not(:first-child) { margin-top: 6px; border-top: 1px solid var(--border); }
  .match-month-select-btn {
    font-family: inherit; font-size: 11.5px; font-weight: 600; cursor: pointer;
    padding: 3px 8px; border-radius: 999px; border: 1px solid var(--border);
    background: var(--page-plane); color: var(--text-secondary);
  }
  .match-month-select-btn:hover { background: var(--grid); }

  .card {
    background: var(--surface-1);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 20px 24px;
    margin-bottom: 20px;
  }
  .card h2 { font-size: 15px; margin: 0 0 16px; color: var(--text-primary); }

  .lineup-court {
    background: var(--page-plane); border: 1px solid var(--border); border-radius: 14px;
    padding: 20px 14px;
  }
  .lineup-row { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; max-width: 460px; margin: 0 auto; }
  .lineup-row + .lineup-row { margin-top: 16px; padding-top: 16px; border-top: 1px dashed var(--border); }
  .jersey-slot { display: flex; flex-direction: column; align-items: center; gap: 6px; min-width: 0; }
  .jersey-name-box {
    position: relative; display: inline-flex; align-items: center; justify-content: center;
    max-width: 100%; box-sizing: border-box; padding: 12px 8px; border-radius: 10px;
    background: var(--surface-1); border: 1px solid var(--border);
    color: var(--text-primary); font-size: 14px; font-weight: 700; text-align: center;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  }
  .jersey-badge {
    position: absolute; right: -6px; top: -6px; width: 22px; height: 22px; border-radius: 50%;
    background: #e5484d; color: #fff; font-size: 11px; font-weight: 700;
    display: flex; align-items: center; justify-content: center;
    border: 2px solid var(--page-plane); box-shadow: 0 1px 2px rgba(0,0,0,0.25);
  }
  @media (max-width: 480px) {
    .lineup-court { padding: 14px 6px; }
    .lineup-row { gap: 4px; }
    .jersey-name-box { font-size: 11.5px; padding: 10px 4px; letter-spacing: -0.2px; }
  }

  .player-grid { display: flex; flex-wrap: wrap; gap: 8px; }
  .player-chip {
    font-family: inherit; font-size: 13.5px; cursor: pointer;
    padding: 8px 16px; border-radius: 20px;
    border: 1px solid var(--border); background: var(--page-plane);
    color: var(--text-primary);
  }
  .player-chip:hover { background: var(--grid); }
  .player-chip.selected { background: var(--series-blue); border-color: var(--series-blue); color: #fff; }
  .player-chip.team-chip { font-weight: 600; border-color: var(--series-orange); color: var(--series-orange); }
  .player-chip.team-chip.selected { background: var(--series-orange); border-color: var(--series-orange); color: #fff; }

  .video-clip-groups { display: flex; flex-wrap: wrap; gap: 10px; }
  .video-clip-list { display: flex; flex-wrap: wrap; gap: 8px; width: 100%; }
  .video-clip-link {
    font-size: 12.5px; text-decoration: none; color: var(--text-primary);
    padding: 5px 10px; border-radius: 14px; border: 1px solid var(--border);
    background: var(--page-plane);
  }
  .video-clip-link:hover { background: var(--grid); }
  .video-clip-link.playing { background: var(--series-blue); border-color: var(--series-blue); color: #fff; }
  .video-player-wrap {
    position: relative; width: 100%; max-width: 520px; aspect-ratio: 16 / 9;
    margin: 4px 0 16px; background: #000; border-radius: 10px; overflow: hidden;
  }
  .video-player-wrap iframe { position: absolute; inset: 0; width: 100%; height: 100%; border: 0; }

  .match-bar { margin-bottom: 18px; }
  .match-bar .label { font-size: 12.5px; color: var(--text-muted); margin-bottom: 8px; }

  .drop-zone-card { margin-bottom: 18px; }
  #dropZoneContent { margin-top: 8px; }
  .drop-zone {
    border: 2px dashed var(--baseline); border-radius: 10px; padding: 20px;
    text-align: center; cursor: pointer; color: var(--text-muted); font-size: 13px;
    transition: background 0.15s, border-color 0.15s, color 0.15s; line-height: 1.7;
  }
  .drop-zone.drag-over { border-color: var(--series-blue); background: var(--page-plane); color: var(--text-primary); }
  .drop-zone-status { margin-top: 10px; }
  .drop-zone-status .hint { margin-top: 4px; }
  .dl-btn {
    display: inline-block; margin: 6px 8px 0 0; padding: 7px 14px; border-radius: 8px;
    background: var(--series-blue); color: #fff; font-size: 12.5px; font-weight: 600;
    text-decoration: none; border: none; cursor: pointer;
  }
  .match-chip.selected { background: var(--series-blue); border-color: var(--series-blue); color: #fff; }

  #detailCard, #detailCard2, #mainArea, #comparisonCard { display: none; }

  .comparison-table td, .comparison-table th { text-align: center; }
  .comparison-table td:first-child, .comparison-table th:first-child {
    text-align: left; position: sticky; left: 0; background: var(--surface-1);
  }
  .comparison-table .section-row td {
    text-align: left; background: var(--grid); font-weight: 600; font-size: 13px;
    color: var(--text-secondary);
  }

  .accordion-section { margin-top: 14px; }
  .accordion-section:first-child { margin-top: 0; }
  .accordion-toggle {
    width: 100%; font-family: inherit; font-size: 15px; font-weight: 600; cursor: pointer;
    padding: 13px 16px; border-radius: 10px; text-align: left;
    border: 1px solid var(--border); background: var(--page-plane); color: var(--text-primary);
    display: flex; align-items: center; justify-content: space-between; gap: 8px;
  }
  .accordion-toggle:hover { background: var(--grid); }
  .accordion-toggle .arrow { color: var(--text-muted); font-size: 12px; flex-shrink: 0; }
  .accordion-body { padding: 16px 2px 4px; }
  .rank-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; }
  .rank-grid h4 { margin: 0 0 8px; font-size: 14px; }

  .page { display: none; }
  .page.active { display: block; }

  .tab-nav { display: flex; gap: 8px; margin-bottom: 20px; }
  .tab-btn {
    font-family: inherit; font-size: 14px; font-weight: 600; cursor: pointer;
    padding: 10px 20px; border-radius: 10px;
    border: 1px solid var(--border); background: var(--surface-1);
    color: var(--text-secondary);
  }
  .tab-btn:hover { background: var(--grid); }
  .tab-btn.active { background: var(--series-blue); border-color: var(--series-blue); color: #fff; }
  .detail-header { display: flex; align-items: baseline; gap: 10px; margin-bottom: 16px; }
  .detail-header h2 { margin: 0; font-size: 18px; }
  .detail-header span { color: var(--text-muted); font-size: 13px; }

  .section-label { font-size: 17px; color: var(--text-primary); margin: 22px 0 10px; font-weight: 600; }

  .collapsible { margin: 22px 0 10px; }
  .collapsible-header {
    display: flex; align-items: center; justify-content: space-between; width: 100%;
    background: none; border: none; padding: 0; margin: 0; cursor: pointer; text-align: left;
    font: inherit; color: inherit;
  }
  .collapsible-header .section-label { margin: 0; }
  .collapsible-header .collapse-arrow {
    color: var(--text-muted); font-size: 13px; margin-left: 10px; flex-shrink: 0;
    transition: transform 0.15s ease;
  }
  .collapsible.closed .collapse-arrow { transform: rotate(-90deg); }
  .collapsible-body { margin-top: 10px; }
  .collapsible.closed .collapsible-body { display: none; }
  .expand-all-btn {
    border: 1px solid var(--border); background: var(--surface-1); color: var(--text-secondary);
    border-radius: 8px; padding: 6px 12px; font-size: 12.5px; cursor: pointer; white-space: nowrap;
  }
  .expand-all-btn:hover { background: var(--grid); }
  .page-toolbar { display: flex; justify-content: flex-end; margin: 0 0 12px; }
  .stat-tiles { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin-bottom: 20px; }
  @media (max-width: 640px) { .stat-tiles { grid-template-columns: repeat(2, 1fr); } }
  .stat-tile { background: var(--page-plane); border: 1px solid var(--border); border-radius: 10px; padding: 12px 14px; }
  .stat-tile .k { font-size: 11.5px; color: var(--text-muted); margin-bottom: 4px; }
  .stat-tile .v { font-size: 20px; font-variant-numeric: tabular-nums; }
  .pct-count { font-size: 11px; font-weight: 400; color: var(--text-muted); margin-left: 3px; white-space: nowrap; }
  .pct-count-block { font-size: 11px; font-weight: 400; color: var(--text-muted); margin-top: 2px; }

  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th, td { text-align: right; padding: 8px 10px; border-bottom: 1px solid var(--grid); font-variant-numeric: tabular-nums; }
  th:first-child, td:first-child { text-align: left; font-variant-numeric: normal; }
  th { color: var(--text-muted); font-weight: 500; font-size: 12px; }
  .hint { color: var(--text-muted); font-size: 12px; margin-top: 10px; }

  .rotation-table th, .rotation-table td { font-size: 15px; padding: 11px 13px; }
  .rotation-table th { font-size: 13px; }

  .score-chart { margin-bottom: 28px; }
  .score-chart:last-child { margin-bottom: 0; }
  .score-chart-title { font-size: 13px; font-weight: 600; color: var(--text-primary); margin-bottom: 8px; }
  .score-chart-legend { display: flex; gap: 18px; margin-top: 6px; }
  .legend-item { display: flex; align-items: center; gap: 6px; font-size: 12.5px; color: var(--text-secondary); }
  .legend-swatch { width: 14px; height: 3px; border-radius: 2px; display: inline-block; }

  .match-title-row { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; }
  .match-title-row h2 { margin: 0; font-size: 20px; color: var(--text-secondary); }
  .summary-btn {
    font-family: inherit; font-size: 13px; font-weight: 600; cursor: pointer;
    padding: 9px 16px; border-radius: 9px; white-space: nowrap;
    border: 1px solid var(--series-blue); background: var(--surface-1); color: var(--series-blue);
    display: inline-flex; align-items: center; gap: 6px;
  }
  .summary-btn:hover { background: var(--series-blue); color: #fff; }

  .summary-overlay {
    display: none; position: fixed; inset: 0; z-index: 100;
    background: rgba(10,10,10,0.62); padding: 24px;
    align-items: flex-start; justify-content: center; overflow-y: auto;
  }
  .summary-modal {
    background: #ffffff; border-radius: 14px; padding: 16px;
    max-width: 780px; width: 100%; margin: 20px 0;
    box-shadow: 0 20px 60px rgba(0,0,0,0.4);
  }
  .summary-modal-actions { display: flex; justify-content: flex-end; gap: 10px; margin-bottom: 14px; }
  .summary-close-btn {
    font-family: inherit; font-size: 13px; font-weight: 600; cursor: pointer;
    padding: 9px 14px; border-radius: 9px; border: 1px solid #d8d8d4; background: #fff; color: #52514e;
  }
  .summary-close-btn:hover { background: #f2f2ef; }
  .summary-scroll { max-height: 82vh; overflow-y: auto; border-radius: 8px; text-align: center; background: #f2f2ef; padding: 12px; }
  #summaryImg { max-width: 100%; height: auto; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.12); }
  .summary-hint { margin: -4px 0 12px; font-size: 12px; color: var(--text-muted); text-align: right; }
  .summary-set-tabs { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 10px; }
  .summary-set-tab {
    font-family: inherit; font-size: 12.5px; font-weight: 600; cursor: pointer;
    padding: 6px 13px; border-radius: 999px; border: 1px solid var(--border);
    background: var(--surface-1); color: var(--text-secondary);
  }
  .summary-set-tab:hover { background: var(--grid); }
  .summary-set-tab.active { background: var(--series-blue); border-color: var(--series-blue); color: #fff; }
</style>
</head>
<body>
<div class="viz-root">
<div class="wrap">
  <header>
    <h1 id="teamTitle"></h1>
    <p>試合を選ぶと、その試合のスタッツが表示されます</p>
  </header>

  <div class="app-layout">
  <div class="sidebar-column">
  <button type="button" id="sidebarToggleBtn" class="sidebar-toggle-btn" title="メニューを閉じる" aria-label="メニューを閉じる">
    <span id="sidebarToggleIcon">◂</span>
  </button>
  <div class="sidebar" id="sidebar">
    <div class="card drop-zone-card">
      <button type="button" class="player-toggle-btn" id="dropZoneToggleBtn" style="margin-top:0">
        <span>.dvwファイルを追加</span>
        <span class="arrow" id="dropZoneToggleArrow">▾</span>
      </button>
      <div id="dropZoneContent" style="display:none">
        <div id="dropZone" class="drop-zone">
          .dvwファイルをここにドラッグ&ドロップ<br>（またはクリックしてファイルを選ぶ）<br>その場ですぐスタッツが見られます
        </div>
        <input type="file" id="dvwFileInput" accept=".dvw" style="display:none">
        <div class="drop-zone-status" id="dropZoneStatus"></div>
      </div>
    </div>

    <div class="card">
      <div class="sidebar-label">試合を選択</div>
      <button type="button" class="player-toggle-btn" id="matchListToggleBtn" style="margin-top:0">
        <span id="matchListToggleLabel">試合</span>
        <span class="arrow" id="matchListToggleArrow">▾</span>
      </button>
      <div class="player-checklist" id="matchChecklist" style="display:none"></div>
      <p class="hint" id="noMatchesHint" style="display:none">まだ試合データがありません。</p>
      <p class="hint" id="matchComparisonHint" style="display:none">2試合以上選ぶと、並べて比較できます。</p>
      <div id="opponentQuickSelectWrap" style="display:none;margin-top:10px">
        <div class="sidebar-label" style="margin-bottom:6px">対戦相手でまとめて選ぶ（スカウティング用）</div>
        <select id="opponentQuickSelect" class="player-toggle-btn" style="margin-top:0;width:100%"></select>
      </div>
    </div>

    <div class="card" id="playerSelectCard" style="display:none">
      <div class="sidebar-label">選手を選択</div>
      <div class="player-grid" id="teamGrid"></div>
      <button type="button" class="player-toggle-btn" id="playerListToggleBtn">
        <span id="playerListToggleLabel">選手</span>
        <span class="arrow" id="playerListToggleArrow">▾</span>
      </button>
      <div class="player-checklist" id="playerChecklist" style="display:none"></div>
    </div>
  </div>
  </div>

  <div class="main-content">
  <div class="card" id="matchComparisonCard" style="display:none">
    <div class="detail-header">
      <h2 id="matchComparisonTitle"></h2>
      <button type="button" class="expand-all-btn" style="margin-left:auto" onclick="toggleAllAccordionSections(this)">すべて展開</button>
    </div>
    <p class="hint" style="margin-top:-6px">試合を2つ以上選ぶと、ここに並べて比較表示されます。もう一度チェックを外すと選択解除できます。</p>
    <div id="matchComparisonBody"></div>
  </div>
  <div id="mainArea">
  <div class="match-title-row">
    <h2 id="matchTitle"></h2>
    <button type="button" class="summary-btn" id="summaryBtn" style="display:none">🖼 試合サマリーを1枚で見る</button>
  </div>

  <div class="tab-nav">
    <button class="tab-btn active" id="tabBtn1" onclick="showPage(1)">選手別スタッツ</button>
    <button class="tab-btn" id="tabBtn2" onclick="showPage(2)">チーム分析</button>
  </div>

  <div class="page active" id="page1">
    <div class="card" id="detailCard">
      <div class="detail-header">
        <h2 id="detailName"></h2>
        <span id="detailNumber"></span>
        <button type="button" class="expand-all-btn" style="margin-left:auto" onclick="toggleAllCollapsibles(document.getElementById('page1'), this)">すべて展開</button>
      </div>

      <div class="collapsible" id="detailSpikeSection" style="margin-top:0">
        <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
          <h3 class="section-label">スパイク（全セット合計）</h3>
          <span class="collapse-arrow">▾</span>
        </button>
        <div class="collapsible-body">
          <div class="stat-tiles" id="detailTiles"></div>
          <table>
            <thead>
              <tr><th id="detailSetHeaderLabel">セット</th><th>打数</th><th>得点</th><th>ミス</th><th>被ブロック</th><th>決定率</th><th>効果率</th></tr>
            </thead>
            <tbody id="detailSetBody"></tbody>
          </table>
        </div>
      </div>

      <div class="collapsible closed" id="detailHardSection" style="display:none">
        <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
          <h3 class="section-label">強打（全セット合計）</h3>
          <span class="collapse-arrow">▾</span>
        </button>
        <div class="collapsible-body">
          <p class="hint" style="margin-top:-6px">しっかり打ち込んだ攻撃（スパイクのうち強打だった分）の内訳です。</p>
          <div class="stat-tiles" id="detailHardTiles"></div>
          <table>
            <thead>
              <tr><th id="detailHardSetHeaderLabel">セット</th><th>打数</th><th>得点</th><th>ミス</th><th>決定率</th></tr>
            </thead>
            <tbody id="detailHardBody"></tbody>
          </table>
        </div>
      </div>

      <div class="collapsible closed" id="detailFeintSection" style="display:none">
        <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
          <h3 class="section-label">フェイント（全セット合計）</h3>
          <span class="collapse-arrow">▾</span>
        </button>
        <div class="collapsible-body">
          <p class="hint" style="margin-top:-6px">相手コートへ軽く落とす攻撃（スパイクのうちフェイントだった分）の内訳です。</p>
          <div class="stat-tiles" id="detailFeintTiles"></div>
          <table>
            <thead>
              <tr><th id="detailFeintSetHeaderLabel">セット</th><th>打数</th><th>得点</th><th>ミス</th><th>決定率</th></tr>
            </thead>
            <tbody id="detailFeintBody"></tbody>
          </table>
        </div>
      </div>

      <div class="collapsible closed" id="detailRotationSection" style="display:none">
        <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
          <h3 class="section-label">ローテーション別（スパイク・全セット合計）</h3>
          <span class="collapse-arrow">▾</span>
        </button>
        <div class="collapsible-body">
          <p class="hint" style="margin-top:-6px">どのローテーションで多く打っているか・決まっているかの目安です。</p>
          <div id="detailRotationBody"></div>
        </div>
      </div>
    </div>

    <div class="card" id="detailCard2">
      <div class="collapsible closed" id="blockSection" style="margin-top:0">
        <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
          <h3 class="section-label" id="blockSectionLabel" style="margin-top:0">ブロック（全セット合計）</h3>
          <span class="collapse-arrow">▾</span>
        </button>
        <div class="collapsible-body">
          <div class="stat-tiles" id="blockTiles"></div>
        </div>
      </div>

      <div class="collapsible closed" id="serveSection">
        <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
          <h3 class="section-label" id="serveSectionLabel">サーブ（全セット合計）</h3>
          <span class="collapse-arrow">▾</span>
        </button>
        <div class="collapsible-body">
          <div class="stat-tiles" id="serveTiles"></div>
        </div>
      </div>

      <div class="collapsible closed" id="receiveSection">
        <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
          <h3 class="section-label" id="receiveSectionLabel">レシーブ（全セット合計）</h3>
          <span class="collapse-arrow">▾</span>
        </button>
        <div class="collapsible-body">
          <div class="stat-tiles" id="receiveTiles"></div>
        </div>
      </div>

      <div class="collapsible closed" id="videoClipsSection" style="display:none">
        <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
          <h3 class="section-label">プレー映像</h3>
          <span class="collapse-arrow">▾</span>
        </button>
        <div class="collapsible-body">
          <p class="hint" style="margin-top:-6px">
            プレーを押すと、その場面からすぐに動画が再生されます（動画側の時刻合わせが多少ずれることがあります）。
          </p>
          <div class="video-player-wrap" id="videoPlayerWrap" style="display:none">
            <iframe id="videoPlayerFrame" src="" allow="autoplay; encrypted-media" allowfullscreen></iframe>
          </div>
          <div id="videoClipsBody"></div>
        </div>
      </div>

      <p class="hint" id="errorBreakdownLine" style="display:none;margin-top:16px;margin-bottom:0"></p>
    </div>

    <div class="card collapsible closed" id="comparisonCard">
      <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
        <div class="detail-header" style="margin-bottom:0">
          <h2 id="comparisonTitle"></h2>
        </div>
        <span class="collapse-arrow">▾</span>
      </button>
      <div class="collapsible-body">
        <p class="hint" style="margin-top:-6px">選手を2人以上選ぶと、ここに並べて比較表示されます。もう一度チップを押すと選択解除できます。</p>
        <div id="comparisonBody"></div>
      </div>
    </div>

    <div class="card collapsible closed" id="startingLineupsCard">
      <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
        <h3 class="section-label" style="margin-top:0">スタメン（セット開始時点）</h3>
        <span class="collapse-arrow">▾</span>
      </button>
      <div class="collapsible-body">
        <p class="hint" style="margin-top:-6px">S1が基本的にサーブ位置（次に回転して前衛に上がる並び）です。セットを選んでください。</p>
        <div class="player-grid" id="startingLineupSetButtons"></div>
        <div id="startingLineupBody" style="margin-top:16px"></div>
      </div>
    </div>

    <div class="card collapsible closed" id="leaderboardCard">
      <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
        <h3 class="section-label" id="leaderboardLabel" style="margin-top:0">選手ランキング（この試合）</h3>
        <span class="collapse-arrow">▾</span>
      </button>
      <div class="collapsible-body">
        <p class="hint" style="margin-top:-6px">
          規定本数（3本）以上の選手が対象です。
        </p>
        <div id="leaderboardBody"></div>
      </div>
    </div>
  </div>

  <div class="page" id="page2">
    <div class="page-toolbar">
      <button type="button" class="expand-all-btn" onclick="toggleAllCollapsibles(document.getElementById('page2'), this)">すべて展開</button>
    </div>

    <div class="card">
      <div class="collapsible" id="sideOutBreakSection" style="margin-top:0">
        <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
          <h3 class="section-label" id="sideOutBreakLabel" style="margin-top:0">サイドアウト率・ブレイク率（チーム全体）</h3>
          <span class="collapse-arrow">▾</span>
        </button>
        <div class="collapsible-body">
          <p class="hint" style="margin-top:-6px" id="sideOutBreakHint">
            サイドアウト率＝相手のサーブを受けたラリーで勝った割合／ブレイク率＝自分たちのサーブのラリーで勝った割合です。セットを選んでください。
          </p>
          <div class="player-grid" id="sideOutBreakSetButtons"></div>
          <div class="stat-tiles" id="sideOutBreakTiles" style="margin-top:16px"></div>
        </div>
      </div>

      <div class="collapsible closed" id="sideOutBreakByRotationSection">
        <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
          <h3 class="section-label" id="sideOutBreakByRotationLabel">ローテーション別（全セット合計）</h3>
          <span class="collapse-arrow">▾</span>
        </button>
        <div class="collapsible-body">
          <p class="hint" style="margin-top:-6px">
            どのローテーションで失点しやすいか・得点しやすいかの目安です。
          </p>
          <div id="sideOutBreakByRotationBody"></div>
          <p class="hint" id="rotationInsightLine" style="margin-top:12px"></p>
        </div>
      </div>
    </div>

    <div class="card collapsible closed" id="scoringPatternCard">
      <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
        <h3 class="section-label" style="margin-top:0">得点パターン（チーム全体）</h3>
        <span class="collapse-arrow">▾</span>
      </button>
      <div class="collapsible-body">
        <p class="hint" style="margin-top:-6px">
          自チームがどうやって得点し、どうやって失点しているかの内訳です。「相手ミス」「自チームミス」は、
          自分たちの決定力・安定感がどれくらい得点に影響しているかの目安になります。
        </p>
        <div class="player-grid" id="scoringPatternMatchButtons" style="margin-bottom:12px"></div>
        <div id="scoringPatternBody"></div>
      </div>
    </div>

    <div class="card" id="scoreProgressionCard">
      <div class="collapsible closed" id="scoreProgressionSection" style="margin-top:0">
        <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
          <h3 class="section-label" style="margin-top:0">セット経過（得点推移）</h3>
          <span class="collapse-arrow">▾</span>
        </button>
        <div class="collapsible-body">
          <p class="hint" style="margin-top:-6px">
            ラリーが決まるたびに、その時点のスコアをグラフにしています。セットを選んでください。
          </p>
          <div class="player-grid" id="scoreProgressionSetButtons"></div>
          <div id="scoreProgressionChartArea" style="margin-top:16px"></div>
        </div>
      </div>

      <div class="collapsible closed" id="runHighlightsSection">
        <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
          <h3 class="section-label" id="runHighlightsLabel">セットごとの流れハイライト</h3>
          <span class="collapse-arrow">▾</span>
        </button>
        <div class="collapsible-body">
          <div class="player-grid" id="runHighlightsMatchButtons" style="margin-bottom:12px"></div>
          <div id="runHighlightsBody"></div>
        </div>
      </div>
    </div>

    <div class="card collapsible closed" id="opponentCourseCard">
      <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
        <h3 class="section-label" style="margin-top:0">相手が決めた攻撃のコース（％）</h3>
        <span class="collapse-arrow">▾</span>
      </button>
      <div class="collapsible-body">
        <p class="hint" style="margin-top:-6px">
          相手が得点した攻撃だけを、コンビの分類でレフト／ミドル／ライト／バックアタックに分けています。
          実際の着地位置ではなく、コースの分類ごとの代表的な位置に矢印を表示した簡易図です。セットを選んでください。
        </p>
        <div class="player-grid" id="opponentCourseSetButtons" style="margin-bottom:12px"></div>
        <div class="player-grid" id="opponentCourseFilterButtons" style="margin-bottom:10px"></div>
        <div id="opponentCourseBody"></div>
        <div id="opponentCourseDiagram" style="margin-top:12px"></div>
      </div>
    </div>

    <div class="card collapsible closed" id="rotationCard">
      <button type="button" class="collapsible-header" onclick="toggleCollapsible(this)">
        <h3 class="section-label" id="rotationCardLabel" style="margin-top:0">ローテーション別 攻撃タイプ分布（チーム全体）</h3>
        <span class="collapse-arrow">▾</span>
      </button>
      <div class="collapsible-body">
        <p class="hint" style="margin-top:-6px" id="rotationCardHint">
          セット開始時のローテーションをS1とし、自チームが1回転するたびにS2→S3…と数えています。
          攻撃タイプ（レフト／ライト／ミドル／バックアタック）は使用した攻撃コンビの分類によるものです。セットを選んでください。
        </p>
        <div class="player-grid" id="rotationSetButtons"></div>
        <h4 style="margin:16px 0 8px;font-size:14px" id="tossDistributionLabel">トス配分（％）</h4>
        <p class="hint" style="margin-top:-4px">
          「その他」はセッターダンプ（ツー）のみが入ります。
        </p>
        <div id="tossDistributionBody"></div>
        <h4 style="margin:16px 0 8px;font-size:14px" id="attackTechniqueLabel">打法（％）</h4>
        <p class="hint" style="margin-top:-4px">
          強打＝しっかり打ち込んだ攻撃、フェイント＝相手コートへ軽く落とす攻撃です。
        </p>
        <div id="attackTechniqueBody"></div>
        <div id="rotationTables" style="margin-top:16px"></div>
      </div>
    </div>
  </div>
  </div>
  </div>
  </div>
</div>
</div>

<div class="summary-overlay" id="summaryOverlay">
  <div class="summary-modal">
    <div class="summary-modal-actions">
      <button type="button" class="summary-btn" id="summarySaveBtn">画像として保存</button>
      <button type="button" class="summary-close-btn" id="summaryCloseBtn">✕ 閉じる</button>
    </div>
    <div class="summary-set-tabs" id="summarySetTabs"></div>
    <p class="summary-hint">保存できない場合は、下の画像を長押しして保存できます</p>
    <div class="summary-scroll">
      <p class="hint" id="summaryLoading">読み込み中…</p>
      <img id="summaryImg" alt="試合サマリー画像" style="display:none">
    </div>
  </div>
</div>

<script>
// ==================== dvwファイルをブラウザだけで解析するエンジン ====================
// Python版(stats.py)の集計ロジックと完全に同じ考え方で書いている。
// 何か数字が合わないときは、まずstats.py側の同じ関数名のコメントを読むと分かりやすい。
// OWN_TEAM_CODE と ROSTER は、実際のHTMLではチームごとの値に置き換えて埋め込む。

const CATEGORY_ORDER = ['left', 'right', 'middle', 'pipe', 'other'];
const CATEGORY_LABELS = { left: 'レフト', right: 'ライト', middle: 'ミドル', pipe: 'バックアタック', other: 'その他' };
const COMBO_CATEGORY_MAP = { F: 'left', B: 'right', C: 'middle', P: 'pipe' };
// VolleyStationのテンプレートの説明文と実際の運用が食い違うコンビコードの個別補正
// （P2はテンプレート上「緊急時の2段トス攻撃」だが、実際は「短いレフトのハイセット」として
// 使われている。こうせいさんの説明により2026-08-29確認）
// P9・L9はコンビ表では分類文字'B'（ライト）扱いだが、実際はゾーン9＝後衛からのバックライト攻撃
// なので「バックアタック」（内部キーは従来のpipeのまま）にまとめる。P3は分類文字が空欄だが
// 実際にはミドル攻撃として使われている（いずれもこうせいさんに確認、2026-08-31）。
const COMBO_CATEGORY_OVERRIDES = { P2: 'left', P9: 'pipe', L9: 'pipe', P3: 'middle' };
// 打法（強打／フェイント）：H=強打、T=フェイント（こうせいさんに確認、2026-08-29）。
// ※このH/Tはスキル直後の1文字（例 "*05AH#P5"の"H"）ではなく、コードの末尾側
// （'~'区切りの最後のかたまり。例 "*05AH#P5~24~H2"の"H2"）に入っている「ショットタイプ」の文字。
// スキル直後の文字は実データを見るとM/H/Q/N/Oで、これは「トスのテンポ」を表しており、
// 強打/フェイントの区別ではなかった（2026-08-29、実データで確認）。
const TECH_ORDER = ['hard', 'feint', 'other'];
const TECH_LABELS = { hard: '強打', feint: 'フェイント', other: 'その他' };
const DV_ATTACK_SHOT_TYPE_RE = /([A-Z])(\d)/;

// アタックの生コード全体（例 '*05AH#P5~24~H2'）から、末尾の「ショットタイプ」文字を取り出す
function dvAttackShotType(code) {
  const parts = code.split('~');
  if (parts.length < 2) return null;
  const m = DV_ATTACK_SHOT_TYPE_RE.exec(parts[parts.length - 1]);
  return m ? m[1] : null;
}

const DV_SCORE_RE = /^([*a])p(\d+):(\d+)/;
const DV_ACTION_RE = /^([*a])(\d\d)([ABSR])([A-Za-z])([#+\-!=/])([A-Z0-9]{2})?/;
// 得点パターン集計専用：ディグ(D)・セット(E)・フリーボール(F)も含めた「全スキル」版。
// ディグミスなどでラリーが終わるケースも取りこぼさずに拾うため、DV_ACTION_REとは別に用意している。
const DV_ACTION_ANY_SKILL_RE = /^([*a])(\d\d)([A-Z])([A-Za-z])([#+\-!=/])([A-Z0-9]{2})?/;

function dvRate(numerator, denominator) {
  if (!denominator) return null;
  return numerator / denominator;
}

function dvEmptyPlayerStats() {
  const stats = {
    attempts: 0, points: 0, errors: 0, blocked: 0,
    back_attempts: 0, back_points: 0, back_errors: 0,
    block_attempts: 0, block_points: 0, block_errors: 0,
    block_touch_own: 0, block_touch_opp: 0,
    serve_attempts: 0, serve_aces: 0, serve_errors: 0, serve_half_credit: 0,
    receive_attempts: 0, receive_a: 0, receive_b: 0,
    receive_c: 0, receive_d: 0, receive_errors: 0,
  };
  // 打法（強打／フェイント／その他）別の選手ごとの内訳（2026-08-29追加）
  TECH_ORDER.forEach(t => {
    stats[`tech_${t}_attempts`] = 0;
    stats[`tech_${t}_points`] = 0;
    stats[`tech_${t}_errors`] = 0;
  });
  return stats;
}

function dvEmptyRotationEntry() {
  const categories = {};
  CATEGORY_ORDER.forEach(cat => { categories[cat] = { attempts: 0, points: 0, errors: 0 }; });
  const techniques = {};
  TECH_ORDER.forEach(t => { techniques[t] = { attempts: 0, points: 0, errors: 0 }; });
  return { attempts: 0, points: 0, errors: 0, categories, techniques };
}

function dvGetSection(content, name) {
  const marker = '[' + name + ']';
  let start = content.indexOf(marker);
  if (start === -1) return '';
  const lineEnd = content.indexOf('\n', start);
  start = (lineEnd === -1) ? content.length : lineEnd + 1;
  let end = content.indexOf('[3', start);
  if (end === -1) end = content.length;
  return content.slice(start, end);
}

function dvGetOwnTeamMarker(content, ownTeamCode) {
  const section = dvGetSection(content, '3TEAMS');
  const lines = section.split('\n').filter(l => l.trim());
  const codes = lines.map(l => l.split(';')[0]);
  return codes[0] === ownTeamCode ? '*' : 'a';
}

function dvGetBackAttackCombos(content) {
  const section = dvGetSection(content, '3ATTACKCOMBINATION');
  const backCombos = new Set();
  section.split('\n').forEach(line => {
    if (!line.trim()) return;
    const fields = line.split(';');
    const code = fields[0];
    const zone = fields.length > 1 ? fields[1] : '';
    if (zone === '7' || zone === '8' || zone === '9') backCombos.add(code);
  });
  return backCombos;
}

function dvGetComboCategories(content) {
  const section = dvGetSection(content, '3ATTACKCOMBINATION');
  const mapping = {};
  section.split('\n').forEach(line => {
    if (!line.trim()) return;
    const fields = line.split(';');
    const code = fields[0];
    const categoryLetter = fields.length > 8 ? fields[8] : '';
    mapping[code] = COMBO_CATEGORY_MAP[categoryLetter] || 'other';
  });
  Object.assign(mapping, COMBO_CATEGORY_OVERRIDES);
  return mapping;
}

function dvGetMatchDate(content) {
  const section = dvGetSection(content, '3MATCH');
  const firstLine = section.split('\n')[0];
  return firstLine.split(';')[0];
}

function dvDateSlug(matchDate) {
  const parts = matchDate.split('/');
  if (parts.length === 3) {
    const [day, month, year] = parts;
    return `${year}-${month}-${day}`;
  }
  return matchDate.replace(/\//g, '-');
}

function dvBuildMatchSlug(matchDate, opponent) {
  return `${dvDateSlug(matchDate)}_${opponent}`;
}

// 保存形式(DD/MM/YYYY)を、画面表示用に年月日順(YYYY/MM/DD)へ変換する
function formatMatchDate(matchDate) {
  const parts = matchDate.split('/');
  if (parts.length === 3) {
    const [day, month, year] = parts;
    return `${year}/${month}/${day}`;
  }
  return matchDate;
}

function dvSplitIntoSets(lines) {
  const sets = [];
  let current = [];
  lines.forEach(line => {
    if (/^\*\*\dset/.test(line)) {
      sets.push(current);
      current = [];
    } else {
      current.push(line);
    }
  });
  if (current.length) sets.push(current);
  return sets;
}

function dvAnalyzeSet(linesInSet, ownMarker, backAttackCombos) {
  const stats = {};
  function getPlayerStats(number) {
    if (!(number in stats)) stats[number] = dvEmptyPlayerStats();
    return stats[number];
  }
  linesInSet.forEach(line => {
    const fields = line.split(';');
    const code = fields[0];
    const m = DV_ACTION_RE.exec(code);
    if (!m) return;
    const team = m[1], playerNum = m[2], skill = m[3], evaluation = m[5], combo = m[6];
    if (team !== ownMarker) return;
    const number = parseInt(playerNum, 10);
    const p = getPlayerStats(number);

    if (skill === 'A') {
      const isBack = !!(combo && backAttackCombos.has(combo));
      p.attempts += 1;
      if (isBack) p.back_attempts += 1;
      if (evaluation === '#') {
        p.points += 1;
        if (isBack) p.back_points += 1;
      } else if (evaluation === '=') {
        p.errors += 1;
        if (isBack) p.back_errors += 1;
      } else if (evaluation === '/') {
        p.blocked += 1;
      }

      // 打法（強打／フェイント／その他）別の内訳も選手ごとに集計する
      const shot = dvAttackShotType(code);
      const tech = shot === 'H' ? 'hard' : (shot === 'T' ? 'feint' : 'other');
      p[`tech_${tech}_attempts`] += 1;
      if (evaluation === '#') p[`tech_${tech}_points`] += 1;
      else if (evaluation === '=') p[`tech_${tech}_errors`] += 1;
    } else if (skill === 'B') {
      p.block_attempts += 1;
      if (evaluation === '#') p.block_points += 1;
      else if (evaluation === '=') p.block_errors += 1;
      else if (evaluation === '/') p.block_errors += 1;
      else if (evaluation === '!') p.block_touch_opp += 1;
      else if (evaluation === '+') p.block_touch_own += 1;
    } else if (skill === 'S') {
      // '+'・'/' は、エースではないが相手のレシーブを崩せている効果的なサーブなので、
      // 効果率に半分の重みで加える。'-'は相手にAパスされている（＝崩せていない）ので対象外
      // （こうせいさんの指定、2026-08-30。以前は'-'を含めていたが変更）
      p.serve_attempts += 1;
      if (evaluation === '#') p.serve_aces += 1;
      else if (evaluation === '=') p.serve_errors += 1;
      else if (evaluation === '+' || evaluation === '/') p.serve_half_credit += 1;
    } else if (skill === 'R') {
      p.receive_attempts += 1;
      if (evaluation === '#') p.receive_a += 1;
      else if (evaluation === '+') p.receive_b += 1;
      else if (evaluation === '=') p.receive_errors += 1;
      else if (evaluation === '!') p.receive_c += 1;
      else if (evaluation === '-' || evaluation === '/') p.receive_d += 1;
      else p.receive_c += 1;
    }
  });
  return stats;
}

function dvAnalyzeRotationAttacks(linesInSet, ownMarker, comboCategories) {
  const rotationStats = {};
  for (let n = 1; n <= 6; n++) rotationStats[n] = dvEmptyRotationEntry();
  let ownRotationKey = null;
  let rotationNumber = 1;

  linesInSet.forEach(line => {
    const fields = line.split(';');
    const code = fields[0];

    if (fields.length >= 26) {
      const slice = ownMarker === '*' ? fields.slice(14, 20) : fields.slice(20, 26);
      const isAllEmpty = slice.every(v => v === '');
      if (!isAllEmpty) {
        const key = JSON.stringify(slice);
        if (ownRotationKey === null) {
          ownRotationKey = key;
        } else if (key !== ownRotationKey) {
          ownRotationKey = key;
          rotationNumber = (rotationNumber % 6) + 1;
        }
      }
    }

    const m = DV_ACTION_RE.exec(code);
    if (!m) return;
    const team = m[1], skill = m[3], skillType = m[4], evaluation = m[5], combo = m[6];
    if (team !== ownMarker || skill !== 'A') return;
    // PPはセッターダンプではなく「ダイレクト」攻撃で、トス配分・打法別・ローテーション別の
    // 集計にはなじまないため丸ごと除外する（選手個人の通常のスパイク統計は従来通り。2026-08-31）
    if (combo === 'PP') return;

    const entry = rotationStats[rotationNumber];
    entry.attempts += 1;
    if (evaluation === '#') entry.points += 1;
    else if (evaluation === '=') entry.errors += 1;

    const cat = comboCategories[combo] || 'other';
    const catEntry = entry.categories[cat];
    catEntry.attempts += 1;
    if (evaluation === '#') catEntry.points += 1;
    else if (evaluation === '=') catEntry.errors += 1;

    // 打法：コード末尾のショットタイプ文字がH=強打、T=フェイント（こうせいさんの説明、2026-08-29）
    const shot = dvAttackShotType(code);
    const tech = shot === 'H' ? 'hard' : (shot === 'T' ? 'feint' : 'other');
    const techEntry = entry.techniques[tech];
    techEntry.attempts += 1;
    if (evaluation === '#') techEntry.points += 1;
    else if (evaluation === '=') techEntry.errors += 1;
  });

  return rotationStats;
}

function dvMergeRotationStats(allSetRotationStats) {
  const total = {};
  for (let n = 1; n <= 6; n++) total[n] = dvEmptyRotationEntry();
  allSetRotationStats.forEach(setStats => {
    for (let n = 1; n <= 6; n++) {
      ['attempts', 'points', 'errors'].forEach(k => { total[n][k] += setStats[n][k]; });
      CATEGORY_ORDER.forEach(cat => {
        ['attempts', 'points', 'errors'].forEach(k => {
          total[n].categories[cat][k] += setStats[n].categories[cat][k];
        });
      });
      TECH_ORDER.forEach(t => {
        ['attempts', 'points', 'errors'].forEach(k => {
          total[n].techniques[t][k] += setStats[n].techniques[t][k];
        });
      });
    }
  });
  return total;
}

// ==================== 選手×ローテーションのアタック集計（誰がどのローテで何本打っているか） ====================
function dvAnalyzePlayerRotationAttacks(linesInSet, ownMarker) {
  const result = {};
  for (let n = 1; n <= 6; n++) result[n] = {};
  let ownRotationKey = null;
  let rotationNumber = 1;

  linesInSet.forEach(line => {
    const fields = line.split(';');
    const code = fields[0];

    if (fields.length >= 26) {
      const slice = ownMarker === '*' ? fields.slice(14, 20) : fields.slice(20, 26);
      const isAllEmpty = slice.every(v => v === '');
      if (!isAllEmpty) {
        const key = JSON.stringify(slice);
        if (ownRotationKey === null) {
          ownRotationKey = key;
        } else if (key !== ownRotationKey) {
          ownRotationKey = key;
          rotationNumber = (rotationNumber % 6) + 1;
        }
      }
    }

    const m = DV_ACTION_RE.exec(code);
    if (!m) return;
    const team = m[1], playerNum = m[2], skill = m[3], evaluation = m[5];
    if (team !== ownMarker || skill !== 'A') return;

    const number = parseInt(playerNum, 10);
    const players = result[rotationNumber];
    if (!players[number]) players[number] = { attempts: 0, points: 0, errors: 0 };
    const p = players[number];
    p.attempts += 1;
    if (evaluation === '#') p.points += 1;
    else if (evaluation === '=') p.errors += 1;
  });

  return result;
}

function dvMergePlayerRotationAttacks(allSetResults) {
  const total = {};
  for (let n = 1; n <= 6; n++) total[n] = {};
  allSetResults.forEach(setResult => {
    for (let n = 1; n <= 6; n++) {
      Object.keys(setResult[n]).forEach(numKey => {
        const number = Number(numKey);
        const s = setResult[n][numKey];
        if (!total[n][number]) total[n][number] = { attempts: 0, points: 0, errors: 0 };
        const t = total[n][number];
        t.attempts += s.attempts; t.points += s.points; t.errors += s.errors;
      });
    }
  });
  return total;
}

function dvBuildPlayerRotationSummary(mergedPlayerRotation) {
  const numbers = new Set();
  for (let n = 1; n <= 6; n++) Object.keys(mergedPlayerRotation[n]).forEach(k => numbers.add(Number(k)));
  const summary = {};
  numbers.forEach(number => {
    const rows = [];
    for (let n = 1; n <= 6; n++) {
      const s = mergedPlayerRotation[n][number] || { attempts: 0, points: 0, errors: 0 };
      rows.push({
        rotation: n,
        attempts: s.attempts, points: s.points, errors: s.errors,
        kill_rate: dvRate(s.points, s.attempts),
      });
    }
    summary[number] = rows;
  });
  return summary;
}

function dvRotationSummary(rotationStats) {
  const rows = [];
  const grand = dvEmptyRotationEntry();
  for (let n = 1; n <= 6; n++) {
    const e = rotationStats[n];
    const row = {
      rotation: n,
      attempts: e.attempts, points: e.points, errors: e.errors,
      kill_rate: dvRate(e.points, e.attempts),
      categories: {},
    };
    CATEGORY_ORDER.forEach(cat => {
      row.categories[cat] = {
        attempts: e.categories[cat].attempts,
        points: e.categories[cat].points,
        errors: e.categories[cat].errors,
        kill_rate: dvRate(e.categories[cat].points, e.categories[cat].attempts),
      };
    });
    row.techniques = {};
    TECH_ORDER.forEach(t => {
      row.techniques[t] = {
        attempts: e.techniques[t].attempts,
        points: e.techniques[t].points,
        errors: e.techniques[t].errors,
        kill_rate: dvRate(e.techniques[t].points, e.techniques[t].attempts),
      };
    });
    rows.push(row);
    grand.attempts += e.attempts; grand.points += e.points; grand.errors += e.errors;
    CATEGORY_ORDER.forEach(cat => {
      ['attempts', 'points', 'errors'].forEach(k => { grand.categories[cat][k] += e.categories[cat][k]; });
    });
    TECH_ORDER.forEach(t => {
      ['attempts', 'points', 'errors'].forEach(k => { grand.techniques[t][k] += e.techniques[t][k]; });
    });
  }
  const totalRow = {
    rotation: '合計',
    attempts: grand.attempts, points: grand.points, errors: grand.errors,
    kill_rate: dvRate(grand.points, grand.attempts),
    categories: {},
  };
  CATEGORY_ORDER.forEach(cat => {
    totalRow.categories[cat] = {
      attempts: grand.categories[cat].attempts,
      points: grand.categories[cat].points,
      errors: grand.categories[cat].errors,
      kill_rate: dvRate(grand.categories[cat].points, grand.categories[cat].attempts),
    };
  });
  totalRow.techniques = {};
  TECH_ORDER.forEach(t => {
    totalRow.techniques[t] = {
      attempts: grand.techniques[t].attempts,
      points: grand.techniques[t].points,
      errors: grand.techniques[t].errors,
      kill_rate: dvRate(grand.techniques[t].points, grand.techniques[t].attempts),
    };
  });
  rows.push(totalRow);
  return rows;
}

function dvAnalyzeSideOutBreak(linesInSet, ownMarker) {
  const result = { serve_total: 0, serve_wins: 0, receive_total: 0, receive_wins: 0 };
  let currentServer = null;
  linesInSet.forEach(line => {
    const code = line.split(';')[0];
    const mAction = DV_ACTION_RE.exec(code);
    if (mAction && mAction[3] === 'S') currentServer = mAction[1];
    const mScore = DV_SCORE_RE.exec(code);
    if (mScore && currentServer !== null) {
      const scorer = mScore[1];
      if (currentServer === ownMarker) {
        result.serve_total += 1;
        if (scorer === ownMarker) result.serve_wins += 1;
      } else {
        result.receive_total += 1;
        if (scorer === ownMarker) result.receive_wins += 1;
      }
    }
  });
  return result;
}

function dvMergeSideOutBreak(allSetResults) {
  const total = { serve_total: 0, serve_wins: 0, receive_total: 0, receive_wins: 0 };
  allSetResults.forEach(r => {
    Object.keys(total).forEach(k => { total[k] += r[k]; });
  });
  return total;
}

function dvAnalyzeSideOutBreakByRotation(linesInSet, ownMarker) {
  const result = {};
  for (let n = 1; n <= 6; n++) result[n] = { serve_total: 0, serve_wins: 0, receive_total: 0, receive_wins: 0 };
  let ownRotationKey = null;
  let rotationNumber = 1;
  let currentServer = null;

  linesInSet.forEach(line => {
    const fields = line.split(';');
    const code = fields[0];

    if (fields.length >= 26) {
      const slice = ownMarker === '*' ? fields.slice(14, 20) : fields.slice(20, 26);
      const isAllEmpty = slice.every(v => v === '');
      if (!isAllEmpty) {
        const key = JSON.stringify(slice);
        if (ownRotationKey === null) {
          ownRotationKey = key;
        } else if (key !== ownRotationKey) {
          ownRotationKey = key;
          rotationNumber = (rotationNumber % 6) + 1;
        }
      }
    }

    const mAction = DV_ACTION_RE.exec(code);
    if (mAction && mAction[3] === 'S') currentServer = mAction[1];

    const mScore = DV_SCORE_RE.exec(code);
    if (mScore && currentServer !== null) {
      const scorer = mScore[1];
      const entry = result[rotationNumber];
      if (currentServer === ownMarker) {
        entry.serve_total += 1;
        if (scorer === ownMarker) entry.serve_wins += 1;
      } else {
        entry.receive_total += 1;
        if (scorer === ownMarker) entry.receive_wins += 1;
      }
    }
  });

  return result;
}

function dvMergeSideOutBreakByRotation(allSetResults) {
  const total = {};
  for (let n = 1; n <= 6; n++) total[n] = { serve_total: 0, serve_wins: 0, receive_total: 0, receive_wins: 0 };
  allSetResults.forEach(setResult => {
    for (let n = 1; n <= 6; n++) {
      Object.keys(total[n]).forEach(k => { total[n][k] += setResult[n][k]; });
    }
  });
  return total;
}

function dvSideOutBreakSummary(r) {
  return {
    serve_total: r.serve_total,
    serve_wins: r.serve_wins,
    break_rate: dvRate(r.serve_wins, r.serve_total),
    receive_total: r.receive_total,
    receive_wins: r.receive_wins,
    side_out_rate: dvRate(r.receive_wins, r.receive_total),
  };
}

// 得点推移に加えて、その得点/失点に自チームのどの選手が絡んだかも記録する
// （得点直前の自チームの決定'#'なら得点者、失点直前の自チームのミス'='なら
// そのミスをした選手、相手のミスでの得点は「相手のミス」として選手なしで扱う）
function dvAnalyzeScoreProgression(linesInSet, ownMarker) {
  const points = [{
    own: 0, opponent: 0, scoringTeam: null, scorerNumber: null, byError: false, skill: null,
    opponentScorerNumber: null,
  }];
  let lastTeam = null, lastNumber = null, lastEval = null, lastSkill = null;

  linesInSet.forEach(line => {
    const code = line.split(';')[0];

    const mAction = DV_ACTION_ANY_SKILL_RE.exec(code);
    if (mAction) {
      const team = mAction[1], playerNum = mAction[2], skill = mAction[3], evaluation = mAction[5];
      if (evaluation === '#' || evaluation === '=') {
        lastTeam = team; lastNumber = parseInt(playerNum, 10); lastEval = evaluation; lastSkill = skill;
      }
    }

    const mScore = DV_SCORE_RE.exec(code);
    if (!mScore) return;
    const homeScore = parseInt(mScore[2], 10);
    const awayScore = parseInt(mScore[3], 10);
    const own = ownMarker === '*' ? homeScore : awayScore;
    const opponent = ownMarker === '*' ? awayScore : homeScore;

    const prev = points[points.length - 1];
    let scoringTeam = null, scorerNumber = null, byError = false, skillUsed = null, opponentScorerNumber = null;
    if (own > prev.own) {
      scoringTeam = 'own';
      if (lastTeam === ownMarker && lastEval === '#') {
        scorerNumber = lastNumber;
        skillUsed = lastSkill;
      } else if (lastTeam !== null && lastTeam !== ownMarker && lastEval === '=') {
        byError = true;
        skillUsed = lastSkill;
      }
    } else if (opponent > prev.opponent) {
      scoringTeam = 'opponent';
      if (lastTeam === ownMarker && lastEval === '=') {
        scorerNumber = lastNumber;
        byError = true;
        skillUsed = lastSkill;
      } else if (lastTeam !== null && lastTeam !== ownMarker && lastEval === '#') {
        skillUsed = lastSkill;
        opponentScorerNumber = lastNumber;
      }
    }
    points.push({ own, opponent, scoringTeam, scorerNumber, byError, skill: skillUsed, opponentScorerNumber });
    lastTeam = null; lastNumber = null; lastEval = null; lastSkill = null;
  });
  return points;
}

// ==================== セット開始時点のスタメン（自チームのP1〜P6の背番号）を取り出す ====================
// ローテーション追跡と同じ、行末のならび(fields[14:20]がホーム、fields[20:26]がアウェイ)を使う。
// そのセットで最初に6人分そろって出てくる行が、そのセットのスタメンにあたる。
function dvGetStartingLineup(linesInSet, ownMarker) {
  for (const line of linesInSet) {
    const fields = line.split(';');
    if (fields.length < 26) continue;
    const ownFields = ownMarker === '*' ? fields.slice(14, 20) : fields.slice(20, 26);
    if (ownFields.some(v => v === '')) continue;
    return ownFields.map(v => parseInt(v, 10));
  }
  return null;
}

// 'HH.MM.SS'形式の時刻表記を、真夜中からの経過秒数に変換する
function dvParseWallclockSeconds(t) {
  const [h, m, s] = t.split('.').map(Number);
  if (Number.isNaN(h) || Number.isNaN(m) || Number.isNaN(s)) return null;
  return h * 3600 + m * 60 + s;
}

// 映像連携用に、自チームのアタック・ブロック・サーブ・レシーブ1本ごとの記録
// （誰が・どの評価で・壁時計の何時何分何秒か）を全部リストにしておく。
// 動画側の秒数への変換は、offsetSeconds が分かってから表示側で行う。
function dvBuildPlayLog(linesInSet, ownMarker) {
  const plays = [];
  linesInSet.forEach(line => {
    const fields = line.split(';');
    const code = fields[0];
    const m = DV_ACTION_RE.exec(code);
    if (!m) return;
    const team = m[1], playerNum = m[2], skill = m[3], evaluation = m[5];
    if (team !== ownMarker) return;
    if (fields.length <= 7 || !fields[7]) return;
    const wallclock = dvParseWallclockSeconds(fields[7]);
    if (wallclock === null) return;
    plays.push({ number: parseInt(playerNum, 10), skill, evaluation, wallclock });
  });
  return plays;
}

function dvMergeStats(allSetStats) {
  const total = {};
  allSetStats.forEach(setStats => {
    Object.keys(setStats).forEach(number => {
      if (!(number in total)) total[number] = dvEmptyPlayerStats();
      const s = setStats[number];
      Object.keys(s).forEach(key => { total[number][key] += s[key]; });
    });
  });
  return total;
}

function dvPlayerSummary(number, s, roster, name) {
  const attempts = s.attempts, points = s.points, errors = s.errors, blocked = s.blocked;
  const result = {
    number: number,
    name: name !== undefined ? name : ((roster && roster[number]) || ('#' + number)),
    attempts, points, errors, blocked,
    kill_rate: dvRate(points, attempts),
    efficiency: dvRate(points - errors - blocked, attempts),
    back_attempts: s.back_attempts, back_points: s.back_points, back_errors: s.back_errors,
    block_attempts: s.block_attempts, block_points: s.block_points, block_errors: s.block_errors,
    block_touch_own: s.block_touch_own, block_touch_opp: s.block_touch_opp,
    serve_attempts: s.serve_attempts, serve_aces: s.serve_aces, serve_errors: s.serve_errors,
    serve_half_credit: s.serve_half_credit,
    serve_plus: s.serve_plus, serve_exclaim: s.serve_exclaim,
    serve_ace_rate: dvRate(s.serve_aces, s.serve_attempts),
    serve_error_rate: dvRate(s.serve_errors, s.serve_attempts),
    serve_efficiency: dvRate(s.serve_aces - s.serve_errors + 0.5 * s.serve_half_credit, s.serve_attempts),
    receive_attempts: s.receive_attempts, receive_a: s.receive_a, receive_b: s.receive_b,
    receive_c: s.receive_c, receive_d: s.receive_d, receive_errors: s.receive_errors,
    receive_return_rate: dvRate(s.receive_attempts - s.receive_errors, s.receive_attempts),
    receive_a_rate: dvRate(s.receive_a, s.receive_attempts),
    receive_ab_rate: dvRate(s.receive_a + s.receive_b, s.receive_attempts),
  };
  TECH_ORDER.forEach(tech => {
    const tAttempts = s[`tech_${tech}_attempts`];
    const tPoints = s[`tech_${tech}_points`];
    const tErrors = s[`tech_${tech}_errors`];
    result[`tech_${tech}_attempts`] = tAttempts;
    result[`tech_${tech}_points`] = tPoints;
    result[`tech_${tech}_errors`] = tErrors;
    result[`tech_${tech}_kill_rate`] = dvRate(tPoints, tAttempts);
  });
  return result;
}

function dvTeamStats(statsDict) {
  const total = dvEmptyPlayerStats();
  Object.values(statsDict).forEach(s => {
    Object.keys(total).forEach(key => { total[key] += s[key]; });
  });
  return total;
}

// content: .dvwファイルの中身(文字列)。opponentName: 対戦相手の名前(手入力)。
// ownTeamCode: '3TEAMS'セクションで自チームを表すコード('TSJ'や'TSD')。
// teamLabel: 表示用のチーム名。roster: {背番号: 名前} の対応表。
function dvBuildDashboardData(content, opponentName, ownTeamCode, teamLabel, roster) {
  const ownMarker = dvGetOwnTeamMarker(content, ownTeamCode);
  const backAttackCombos = dvGetBackAttackCombos(content);
  const comboCategories = dvGetComboCategories(content);
  const matchDate = dvGetMatchDate(content);

  const scoutSection = dvGetSection(content, '3SCOUT');
  const allLines = scoutSection.split('\n').filter(l => l.trim());
  const setsLines = dvSplitIntoSets(allLines);

  const allSetStats = setsLines.map(s => dvAnalyzeSet(s, ownMarker, backAttackCombos));
  const totalStats = dvMergeStats(allSetStats);
  const allRotationStats = setsLines.map(s => dvAnalyzeRotationAttacks(s, ownMarker, comboCategories));
  const allSideOutBreak = setsLines.map(s => dvAnalyzeSideOutBreak(s, ownMarker));
  const allSideOutBreakByRotation = setsLines.map(s => dvAnalyzeSideOutBreakByRotation(s, ownMarker));
  // 相手チーム分析（スカウティング）用：ownとoppを入れ替えて同じ関数を再利用するだけで、
  // 相手チーム自身のローテーション番号でのサイドアウト率・ブレイク率が計算できる
  const opponentMarker = ownMarker === '*' ? 'a' : '*';
  const allOppSideOutBreakByRotation = setsLines.map(s => dvAnalyzeSideOutBreakByRotation(s, opponentMarker));
  // 相手が決めた攻撃のコース（レフト/ミドル/ライト）用：同じくownとoppを入れ替えるだけで、
  // 相手自身のコンビ分類の攻撃タイプ内訳が計算できる（2026-09-03追加）
  const allOppRotationStats = setsLines.map(s => dvAnalyzeRotationAttacks(s, opponentMarker, comboCategories));
  const allPlayerRotationAttacks = setsLines.map(s => dvAnalyzePlayerRotationAttacks(s, ownMarker));
  const allScoreProgression = setsLines.map(s => dvAnalyzeScoreProgression(s, ownMarker));
  const allStartingLineups = setsLines.map(s => dvGetStartingLineup(s, ownMarker));
  const allPlayLogs = setsLines.map(s => dvBuildPlayLog(s, ownMarker));

  const numbers = Object.keys(totalStats).map(Number).sort((a, b) => a - b);

  return {
    matchDate: matchDate,
    opponent: opponentName,
    teamLabel: teamLabel,
    sets: allSetStats.map((setStats, i) => ({
      setNumber: i + 1,
      players: Object.keys(setStats).map(Number).sort((a, b) => a - b)
        .map(n => dvPlayerSummary(n, setStats[n], roster)),
    })),
    total: numbers.map(n => dvPlayerSummary(n, totalStats[n], roster)),
    team: {
      name: 'チーム全体',
      total: dvPlayerSummary(null, dvTeamStats(totalStats), roster, 'チーム全体'),
      sets: allSetStats.map((setStats, i) => ({
        setNumber: i + 1,
        ...dvPlayerSummary(null, dvTeamStats(setStats), roster, 'チーム全体'),
      })),
    },
    rotation: {
      bySet: allRotationStats.map((rs, i) => ({ setNumber: i + 1, rows: dvRotationSummary(rs) })),
      total: { rows: dvRotationSummary(dvMergeRotationStats(allRotationStats)) },
      byPlayer: dvBuildPlayerRotationSummary(dvMergePlayerRotationAttacks(allPlayerRotationAttacks)),
    },
    sideOutBreak: {
      bySet: allSideOutBreak.map((r, i) => ({ setNumber: i + 1, ...dvSideOutBreakSummary(r) })),
      total: dvSideOutBreakSummary(dvMergeSideOutBreak(allSideOutBreak)),
      byRotation: (() => {
        const merged = dvMergeSideOutBreakByRotation(allSideOutBreakByRotation);
        const rows = [];
        for (let n = 1; n <= 6; n++) rows.push({ rotation: n, ...dvSideOutBreakSummary(merged[n]) });
        return rows;
      })(),
    },
    opponentSideOutBreak: {
      byRotation: (() => {
        const merged = dvMergeSideOutBreakByRotation(allOppSideOutBreakByRotation);
        const rows = [];
        for (let n = 1; n <= 6; n++) rows.push({ rotation: n, ...dvSideOutBreakSummary(merged[n]) });
        return rows;
      })(),
    },
    opponentRotation: {
      bySet: allOppRotationStats.map((rs, i) => ({ setNumber: i + 1, rows: dvRotationSummary(rs) })),
      total: { rows: dvRotationSummary(dvMergeRotationStats(allOppRotationStats)) },
    },
    scoreProgression: {
      bySet: allScoreProgression.map((points, i) => ({
        setNumber: i + 1, points, startingLineup: allStartingLineups[i],
      })),
    },
    // ブラウザに直接ドラッグ&ドロップした試合には、まだ動画URL・時刻合わせ情報が
    // ないので空にしておく（Python側で処理した試合には入ることがある）。
    video: {
      url: '',
      offsetSeconds: null,
      plays: allPlayLogs.flatMap((setPlays, i) => setPlays.map(p => ({ setNumber: i + 1, ...p }))),
    },
  };
}


// ==================== ここから：チームごとの設定 + ドラッグ&ドロップUI ====================
const OWN_TEAM_CODE = '__OWN_TEAM_CODE__';
const ROSTER = __ROSTER_JSON__;

let MATCHES = [];

function dvDownloadJson(obj, filename) {
  const blob = new Blob([JSON.stringify(obj)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function dvUpdateManifestList(matches, matchDate, opponent, dataFile) {
  const filtered = matches.filter(m => !(m.matchDate === matchDate && m.opponent === opponent));
  filtered.push({ matchDate, opponent, dataFile });
  filtered.sort((a, b) => dvDateSlug(b.matchDate) < dvDateSlug(a.matchDate) ? -1 : 1);
  return filtered;
}

function handleDvwFile(file) {
  const statusEl = document.getElementById('dropZoneStatus');
  statusEl.innerHTML = '<p class="hint">読み込み中…</p>';
  const reader = new FileReader();
  reader.onload = () => {
    const content = reader.result;
    let opponent = window.prompt('対戦相手のチーム名を入力してください（.dvwファイルの中では読み取れないため）');
    if (!opponent) {
      statusEl.innerHTML = '<p class="hint">対戦相手の名前が入力されなかったので、中止しました。</p>';
      return;
    }
    opponent = opponent.trim();
    let data;
    try {
      data = dvBuildDashboardData(content, opponent, OWN_TEAM_CODE, document.getElementById('teamTitle').textContent, ROSTER);
    } catch (e) {
      statusEl.innerHTML = '<p class="hint">読み込みに失敗しました。.dvwファイルの中身を確認してください。（' + escapeHtml(String(e)) + '）</p>';
      return;
    }

    // その場ですぐ表示する（プレビューなので、試合一覧のチェック・選手選択はいったん全部外す）
    selectedMatchKeys.clear();
    document.querySelectorAll('#matchChecklist input[type="checkbox"]').forEach(cb => { cb.checked = false; });
    updateMatchListToggleLabel();
    document.getElementById('matchComparisonCard').style.display = 'none';
    selectedPlayerNumbers.clear();
    document.getElementById('playerSelectCard').style.display = 'block';
    renderTeamGrid();
    renderPlayerGrid();
    showSingleMatch(data);

    // GitHubに保存するためのダウンロードを用意する
    const slug = dvBuildMatchSlug(DATA.matchDate, DATA.opponent);
    const dataFile = slug + '.json';
    const updatedMatches = dvUpdateManifestList(MATCHES, DATA.matchDate, DATA.opponent, dataFile);

    statusEl.innerHTML =
      '<p class="hint">スタッツを表示しました。他のパソコン・スマホでも見られるようにするには、'
      + '下の2つのファイルをダウンロードして、GitHubにアップロードしてください（前と同じ手順です）。</p>'
      + '<button class="dl-btn" id="dlMatchBtn">① ' + escapeHtml(dataFile) + ' をダウンロード</button>'
      + '<button class="dl-btn" id="dlManifestBtn">② matches.json をダウンロード</button>';
    document.getElementById('dlMatchBtn').addEventListener('click', () => dvDownloadJson(DATA, dataFile));
    document.getElementById('dlManifestBtn').addEventListener('click', () => dvDownloadJson(updatedMatches, 'matches.json'));
  };
  reader.onerror = () => {
    statusEl.innerHTML = '<p class="hint">ファイルの読み込みに失敗しました。</p>';
  };
  reader.readAsText(file, 'utf-8');
}

function setupDropZoneToggle() {
  const btn = document.getElementById('dropZoneToggleBtn');
  const content = document.getElementById('dropZoneContent');
  const arrow = document.getElementById('dropZoneToggleArrow');
  btn.addEventListener('click', () => {
    const isOpen = content.style.display !== 'none';
    content.style.display = isOpen ? 'none' : 'block';
    arrow.textContent = isOpen ? '▾' : '▴';
  });
}

function setupDropZone() {
  const zone = document.getElementById('dropZone');
  const input = document.getElementById('dvwFileInput');
  zone.addEventListener('click', () => input.click());
  input.addEventListener('change', () => {
    if (input.files && input.files[0]) handleDvwFile(input.files[0]);
    input.value = '';
  });
  ['dragenter', 'dragover'].forEach(evt => {
    zone.addEventListener(evt, e => {
      e.preventDefault();
      zone.classList.add('drag-over');
    });
  });
  ['dragleave', 'drop'].forEach(evt => {
    zone.addEventListener(evt, e => {
      e.preventDefault();
      zone.classList.remove('drag-over');
    });
  });
  zone.addEventListener('drop', e => {
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) handleDvwFile(file);
  });
}
// ==================== ここまで：チームごとの設定 + ドラッグ&ドロップUI ====================


let DATA = null;

function pct(v) {
  if (v === null || v === undefined) return '-';
  return (v * 100).toFixed(1) + '%';
}

// ％の後ろに「【本数/分母】」を小さめの文字で添える共通表示（例: 40.0%【88/220本】）
function pctCount(v, num, den, unit) {
  unit = unit === undefined ? '本' : unit;
  const countText = (den || den === 0) ? `${num}/${den}${unit}` : `${num}${unit}`;
  return `${pct(v)}<span class="pct-count">【${countText}】</span>`;
}

// pctCountと同じだが、「【本数/分母】」を％の下に改行して表示する（選手ランキングなど、横幅が狭い表向け）
function pctCountBlock(v, num, den, unit) {
  unit = unit === undefined ? '本' : unit;
  const countText = (den || den === 0) ? `${num}/${den}${unit}` : `${num}${unit}`;
  return `${pct(v)}<div class="pct-count-block">【${countText}】</div>`;
}

function renderTeamGrid() {
  const grid = document.getElementById('teamGrid');
  grid.innerHTML = '';
  const chip = document.createElement('button');
  chip.className = 'player-chip team-chip';
  chip.textContent = '全体';
  chip.addEventListener('click', () => selectTeam(chip));
  grid.appendChild(chip);
}

function renderPlayerGrid() {
  const list = document.getElementById('playerChecklist');
  list.innerHTML = '';
  const numbers = Object.keys(ROSTER).map(Number).sort((a, b) => a - b);
  numbers.forEach(number => {
    const label = document.createElement('label');
    label.className = 'player-check-item';
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = selectedPlayerNumbers.has(number);
    cb.addEventListener('change', () => togglePlayerCheckbox(number, cb));
    const span = document.createElement('span');
    span.textContent = ROSTER[number];
    label.appendChild(cb);
    label.appendChild(span);
    list.appendChild(label);
  });
  updatePlayerListToggleLabel();
}

function updatePlayerListToggleLabel() {
  const n = selectedPlayerNumbers.size;
  document.getElementById('playerListToggleLabel').textContent =
    n > 0 ? `選手（${n}人選択中）` : '選手';
}

function togglePlayerCheckbox(number, cbEl) {
  document.querySelectorAll('#teamGrid .player-chip').forEach(c => c.classList.remove('selected'));
  if (cbEl.checked) {
    selectedPlayerNumbers.add(number);
  } else {
    selectedPlayerNumbers.delete(number);
  }
  updatePlayerListToggleLabel();
  renderPlayerSelection();
}

function setupPlayerListToggle() {
  const btn = document.getElementById('playerListToggleBtn');
  const list = document.getElementById('playerChecklist');
  const arrow = document.getElementById('playerListToggleArrow');
  btn.addEventListener('click', () => {
    const isOpen = list.style.display !== 'none';
    list.style.display = isOpen ? 'none' : 'flex';
    arrow.textContent = isOpen ? '▾' : '▴';
  });
}

function setupSidebarToggle() {
  const btn = document.getElementById('sidebarToggleBtn');
  const sidebar = document.getElementById('sidebar');
  const icon = document.getElementById('sidebarToggleIcon');
  btn.addEventListener('click', () => {
    const collapsed = sidebar.classList.toggle('collapsed');
    const label = collapsed ? 'メニューを開く' : 'メニューを閉じる';
    btn.title = label;
    btn.setAttribute('aria-label', label);
    icon.textContent = collapsed ? '▸' : '◂';
  });
}

// 個人／チームのアタックを、打法（強打／フェイント）ごとに「アタック」と同じ形
// （タイル＋セット/試合ごとの内訳表）で表示する（2026-08-29追加）
function renderTechSection(tech, headerElId, tilesElId, bodyElId, totalStats, setRows, hasMatchLabels) {
  const attempts = totalStats[`tech_${tech}_attempts`] || 0;
  const points = totalStats[`tech_${tech}_points`] || 0;
  const errors = totalStats[`tech_${tech}_errors`] || 0;
  const killRate = totalStats[`tech_${tech}_kill_rate`];

  document.getElementById(tilesElId).innerHTML = `
    <div class="stat-tile"><div class="k">打数</div><div class="v">${attempts}</div></div>
    <div class="stat-tile"><div class="k">得点</div><div class="v">${points}</div></div>
    <div class="stat-tile"><div class="k">ミス</div><div class="v">${errors}</div></div>
    <div class="stat-tile"><div class="k">決定率</div><div class="v">${pct(killRate)}</div></div>
  `;

  document.getElementById(headerElId).textContent = hasMatchLabels ? '試合' : 'セット';

  const body = document.getElementById(bodyElId);
  body.innerHTML = '';
  setRows.forEach(sp => {
    const a = sp && sp[`tech_${tech}_attempts`];
    if (sp && a > 0) {
      const tr = document.createElement('tr');
      const rowLabel = sp.matchLabel ? escapeHtml(sp.matchLabel) : `第${sp.setNumber}セット`;
      const p = sp[`tech_${tech}_points`] || 0;
      const e = sp[`tech_${tech}_errors`] || 0;
      const kr = sp[`tech_${tech}_kill_rate`];
      tr.innerHTML = `<td>${rowLabel}</td><td>${a}</td><td>${p}</td><td>${e}</td><td>${pct(kr)}</td>`;
      body.appendChild(tr);
    }
  });
  if (!body.children.length) {
    body.innerHTML = '<tr><td colspan="5" class="hint">出場記録がありません</td></tr>';
  }
}

function markSelected(chipEl) {
  document.querySelectorAll('.player-chip').forEach(c => c.classList.remove('selected'));
  chipEl.classList.add('selected');
}

function showDetail(name, subLabel, totalStats, setRows, rotationRows) {
  document.getElementById('detailCard').style.display = 'block';
  document.getElementById('detailCard2').style.display = 'block';
  document.getElementById('detailName').textContent = name;
  document.getElementById('detailNumber').textContent = subLabel;

  document.getElementById('detailTiles').innerHTML = `
    <div class="stat-tile"><div class="k">打数</div><div class="v">${totalStats.attempts}</div></div>
    <div class="stat-tile"><div class="k">得点</div><div class="v">${totalStats.points}</div></div>
    <div class="stat-tile"><div class="k">ミス</div><div class="v">${totalStats.errors}</div></div>
    <div class="stat-tile"><div class="k">被ブロック</div><div class="v">${totalStats.blocked}</div></div>
    <div class="stat-tile"><div class="k">決定率</div><div class="v">${pct(totalStats.kill_rate)}</div></div>
    <div class="stat-tile"><div class="k">効果率</div><div class="v">${pct(totalStats.efficiency)}</div></div>
  `;

  const hasMatchLabels = setRows.some(sp => sp && sp.matchLabel);
  document.getElementById('detailSetHeaderLabel').textContent = hasMatchLabels ? '試合' : 'セット';

  const setBody = document.getElementById('detailSetBody');
  setBody.innerHTML = '';
  setRows.forEach(sp => {
    if (sp && sp.attempts > 0) {
      const tr = document.createElement('tr');
      const rowLabel = sp.matchLabel ? escapeHtml(sp.matchLabel) : `第${sp.setNumber}セット`;
      tr.innerHTML = `<td>${rowLabel}</td><td>${sp.attempts}</td><td>${sp.points}</td><td>${sp.errors}</td><td>${sp.blocked}</td><td>${pct(sp.kill_rate)}</td>`;
      setBody.appendChild(tr);
    }
  });
  if (!setBody.children.length) {
    setBody.innerHTML = '<tr><td colspan="7" class="hint">出場記録がありません</td></tr>';
  }

  const hardSection = document.getElementById('detailHardSection');
  const feintSection = document.getElementById('detailFeintSection');
  if (totalStats.attempts > 0) {
    hardSection.style.display = '';
    feintSection.style.display = '';
    renderTechSection('hard', 'detailHardSetHeaderLabel', 'detailHardTiles', 'detailHardBody', totalStats, setRows, hasMatchLabels);
    renderTechSection('feint', 'detailFeintSetHeaderLabel', 'detailFeintTiles', 'detailFeintBody', totalStats, setRows, hasMatchLabels);
  } else {
    hardSection.style.display = 'none';
    feintSection.style.display = 'none';
  }

  const rotationSection = document.getElementById('detailRotationSection');
  if (rotationRows && rotationRows.length) {
    rotationSection.style.display = '';
    document.getElementById('detailRotationBody').innerHTML = playerRotationTableHtml(rotationRows);
  } else {
    rotationSection.style.display = 'none';
  }

  const blockSectionWrap = document.getElementById('blockSection');
  const blockTiles = document.getElementById('blockTiles');
  if (totalStats.block_attempts > 0) {
    blockSectionWrap.style.display = '';
    const touch = totalStats.block_touch_own + totalStats.block_touch_opp;
    blockTiles.innerHTML = `
      <div class="stat-tile"><div class="k">本数</div><div class="v">${totalStats.block_attempts}</div></div>
      <div class="stat-tile"><div class="k">得点</div><div class="v">${totalStats.block_points}</div></div>
      <div class="stat-tile"><div class="k">ワンチ</div><div class="v">${touch}</div></div>
      <div class="stat-tile"><div class="k">自コート／相手コート</div><div class="v" style="font-size:16px">${totalStats.block_touch_own} / ${totalStats.block_touch_opp}</div></div>
      <div class="stat-tile"><div class="k">失点</div><div class="v">${totalStats.block_errors}</div></div>
    `;
  } else {
    blockSectionWrap.style.display = 'none';
  }

  const serveSectionWrap = document.getElementById('serveSection');
  const serveTiles = document.getElementById('serveTiles');
  if (totalStats.serve_attempts > 0) {
    serveSectionWrap.style.display = '';
    serveTiles.innerHTML = `
      <div class="stat-tile"><div class="k">打数</div><div class="v">${totalStats.serve_attempts}</div></div>
      <div class="stat-tile"><div class="k">エース</div><div class="v">${totalStats.serve_aces}</div></div>
      <div class="stat-tile"><div class="k">ミス</div><div class="v">${totalStats.serve_errors}</div></div>
      <div class="stat-tile"><div class="k">効果本数</div><div class="v">${totalStats.serve_half_credit}<div style="font-size:11px;color:var(--text-muted);font-weight:400">＋と／の合計</div></div></div>
      <div class="stat-tile"><div class="k">エース率</div><div class="v">${pct(totalStats.serve_ace_rate)}</div></div>
      <div class="stat-tile"><div class="k">ミス率</div><div class="v">${pct(totalStats.serve_error_rate)}</div></div>
      <div class="stat-tile"><div class="k">効果率</div><div class="v">${pct(totalStats.serve_efficiency)}</div></div>
    `;
  } else {
    serveSectionWrap.style.display = 'none';
  }

  const receiveSectionWrap = document.getElementById('receiveSection');
  const receiveTiles = document.getElementById('receiveTiles');
  if (totalStats.receive_attempts > 0) {
    receiveSectionWrap.style.display = '';
    receiveTiles.innerHTML = `
      <div class="stat-tile"><div class="k">本数</div><div class="v">${totalStats.receive_attempts}</div></div>
      <div class="stat-tile"><div class="k">Aパス</div><div class="v">${totalStats.receive_a}</div></div>
      <div class="stat-tile"><div class="k">Bパス</div><div class="v">${totalStats.receive_b}</div></div>
      <div class="stat-tile"><div class="k">Cパス</div><div class="v">${totalStats.receive_c}</div></div>
      <div class="stat-tile"><div class="k">Dパス</div><div class="v">${totalStats.receive_d}</div></div>
      <div class="stat-tile"><div class="k">ミス</div><div class="v">${totalStats.receive_errors}</div></div>
      <div class="stat-tile"><div class="k">返球率</div><div class="v">${pct(totalStats.receive_return_rate)}</div></div>
      <div class="stat-tile"><div class="k">A率</div><div class="v">${pct(totalStats.receive_a_rate)}</div></div>
      <div class="stat-tile"><div class="k">AB率</div><div class="v">${pct(totalStats.receive_ab_rate)}</div></div>
    `;
  } else {
    receiveSectionWrap.style.display = 'none';
  }

  // プレー映像（動画URLが設定されている試合だけ表示される）。選手を切り替えたら、
  // 前の選手の動画が再生されたままにならないように、プレーヤーは毎回リセットする
  const videoSection = document.getElementById('videoClipsSection');
  const videoBody = videoClipsHtml(totalStats.number);
  const videoPlayerWrap = document.getElementById('videoPlayerWrap');
  const videoPlayerFrame = document.getElementById('videoPlayerFrame');
  if (videoPlayerWrap) videoPlayerWrap.style.display = 'none';
  if (videoPlayerFrame) videoPlayerFrame.src = '';
  if (videoBody) {
    videoSection.style.display = '';
    document.getElementById('videoClipsBody').innerHTML = videoBody;
  } else {
    videoSection.style.display = 'none';
    document.getElementById('videoClipsBody').innerHTML = '';
  }

  // ミスの内訳（控えめに一行だけ。派手な表やグラフにはしない）
  const errBreakdownEl = document.getElementById('errorBreakdownLine');
  const errParts = [
    ['スパイク', totalStats.errors],
    ['サーブ', totalStats.serve_errors],
    ['レシーブ', totalStats.receive_errors],
    ['ブロック', totalStats.block_errors],
  ].filter(([, v]) => v);
  if (errParts.length) {
    errBreakdownEl.style.display = '';
    errBreakdownEl.textContent = 'ミスの内訳： ' + errParts.map(([k, v]) => `${k}${v}`).join('／');
  } else {
    errBreakdownEl.style.display = 'none';
  }
}

let selectedPlayerNumbers = new Set();

function renderSingleMatchPlayerView() {
  const nums = Array.from(selectedPlayerNumbers);
  if (nums.length === 0) {
    document.getElementById('detailCard').style.display = 'none';
    document.getElementById('detailCard2').style.display = 'none';
    document.getElementById('comparisonCard').style.display = 'none';
  } else if (selectedMatchKeys.size >= 2) {
    renderCombinedPlayerView(nums);
  } else if (nums.length === 1) {
    document.getElementById('comparisonCard').style.display = 'none';
    const number = nums[0];
    const p = DATA.total.find(pp => pp.number === number);
    if (!p) {
      document.getElementById('detailCard').style.display = 'block';
      document.getElementById('detailCard2').style.display = 'none';
      document.getElementById('detailName').textContent = ROSTER[number] || `#${number}`;
      document.getElementById('detailNumber').textContent = `#${number} ／ この試合の出場記録がありません`;
      document.getElementById('detailTiles').innerHTML = '';
      document.getElementById('detailSetBody').innerHTML =
        '<tr><td colspan="7" class="hint">出場記録がありません</td></tr>';
      document.getElementById('detailHardSection').style.display = 'none';
      document.getElementById('detailFeintSection').style.display = 'none';
      document.getElementById('detailRotationSection').style.display = 'none';
      return;
    }
    const setRows = DATA.sets.map(s => {
      const sp = s.players.find(pp => pp.number === number);
      return sp ? { ...sp, setNumber: s.setNumber } : null;
    });
    const rotationRows = (DATA.rotation.byPlayer && DATA.rotation.byPlayer[number]) || null;
    showDetail(p.name, `#${p.number} ／ 全セット合計`, p, setRows, rotationRows);
  } else {
    document.getElementById('detailCard').style.display = 'none';
    document.getElementById('detailCard2').style.display = 'none';
    showComparison(nums);
  }
}

// ==================== 「選手別スタッツ」タブを、選んだ試合分の合計で表示する ====================
// 試合を2つ以上選んでいるときは、DATA(先頭の試合だけ)ではなく、選んだ試合すべての
// 合計を計算して表示する。試合の選択が変わるたびに呼び直され、その時点で選ばれている
// 試合の分だけを合計し直す。
function renderCombinedPlayerView(nums) {
  const entries = Array.from(selectedMatchKeys)
    .map(k => MATCHES.find(m => matchKey(m) === k))
    .filter(Boolean);
  if (!entries.length) return;
  Promise.all(entries.map(fetchMatchData)).then(dataList => {
    // 非同期で待っている間に選手選択が変わっていたら、古い結果は描画しない
    const stillSame = Array.from(selectedPlayerNumbers).sort().join(',') === nums.slice().sort().join(',');
    if (!stillSame || selectedMatchKeys.size < 2) return;

    if (nums.length === 1) {
      const number = nums[0];
      document.getElementById('comparisonCard').style.display = 'none';
      const perMatchStats = dataList.map(d => (d.total || []).find(p => p.number === number) || null);
      const combined = combineStatsList(perMatchStats);
      const name = ROSTER[number] || `#${number}`;
      if (!combined) {
        document.getElementById('detailCard').style.display = 'block';
        document.getElementById('detailCard2').style.display = 'none';
        document.getElementById('detailName').textContent = name;
        document.getElementById('detailNumber').textContent = `#${number} ／ 選んだ${entries.length}試合に出場記録がありません`;
        document.getElementById('detailTiles').innerHTML = '';
        document.getElementById('detailSetBody').innerHTML =
          '<tr><td colspan="7" class="hint">出場記録がありません</td></tr>';
        document.getElementById('detailHardSection').style.display = 'none';
        document.getElementById('detailFeintSection').style.display = 'none';
        document.getElementById('detailRotationSection').style.display = 'none';
        return;
      }
      const rows = entries.map((entry, i) => {
        const sp = perMatchStats[i];
        return sp ? { ...sp, matchLabel: `${formatMatchDate(entry.matchDate)} ${entry.opponent}` } : null;
      });
      const rotationRowsList = dataList.map(d => (d.rotation.byPlayer && d.rotation.byPlayer[number]) || null);
      const combinedRotationRows = combinePlayerRotationRows(rotationRowsList);
      showDetail(name, `#${number} ／ 選んだ${entries.length}試合の合計`, combined, rows, combinedRotationRows);
    } else {
      document.getElementById('detailCard').style.display = 'none';
      document.getElementById('detailCard2').style.display = 'none';
      showCombinedComparison(nums, dataList, entries.length);
    }
  });
}

function showCombinedComparison(nums, dataList, matchCount) {
  const players = nums.map(n => {
    const perMatchStats = dataList.map(d => (d.total || []).find(p => p.number === n) || null);
    const combined = combineStatsList(perMatchStats);
    if (!combined) return null;
    return { ...combined, number: n, name: ROSTER[n] || `#${n}` };
  }).filter(Boolean);
  document.getElementById('comparisonCard').style.display = 'block';
  document.getElementById('comparisonTitle').textContent = `選手比較（${players.length}人／選んだ${matchCount}試合の合計）`;
  document.getElementById('comparisonBody').innerHTML = comparisonTableHtml(players);
}

function refreshMatchComparisonIfNeeded() {
  if (selectedMatchKeys.size < 2) return;
  const entries = Array.from(selectedMatchKeys)
    .map(k => MATCHES.find(m => matchKey(m) === k))
    .filter(Boolean);
  Promise.all(entries.map(fetchMatchData)).then(dataList => showMatchComparison(dataList));
}

function renderPlayerSelection() {
  // 「選手別スタッツ」タブは、複数試合を選んでいても常に(先頭の試合について)残す
  renderSingleMatchPlayerView();
  // 2試合以上選んでいれば、試合比較カードの選手内訳もあわせて更新する
  refreshMatchComparisonIfNeeded();
}

function comparisonRowHtml(label, players, getter, isPct) {
  let html = `<tr><td>${escapeHtml(label)}</td>`;
  players.forEach(p => {
    const v = getter(p);
    html += `<td>${isPct ? pct(v) : v}</td>`;
  });
  html += '</tr>';
  return html;
}

function comparisonSectionHtml(title, colspan) {
  return `<tr class="section-row"><td colspan="${colspan}">${escapeHtml(title)}</td></tr>`;
}

function comparisonTableHtml(players) {
  const colspan = players.length + 1;
  const hasBlock = players.some(p => p.block_attempts > 0);
  const hasServe = players.some(p => p.serve_attempts > 0);
  const hasReceive = players.some(p => p.receive_attempts > 0);

  let html = '<div style="overflow-x:auto"><table class="rotation-table comparison-table">';
  html += '<thead><tr><th>項目</th>';
  players.forEach(p => {
    html += `<th>${escapeHtml(p.name)}<br><span style="font-weight:400;font-size:12px">#${p.number}</span></th>`;
  });
  html += '</tr></thead><tbody>';

  html += comparisonSectionHtml('スパイク（全セット合計）', colspan);
  html += comparisonRowHtml('打数', players, p => p.attempts, false);
  html += comparisonRowHtml('得点', players, p => p.points, false);
  html += comparisonRowHtml('ミス', players, p => p.errors, false);
  html += comparisonRowHtml('被ブロック', players, p => p.blocked, false);
  html += comparisonRowHtml('決定率', players, p => p.kill_rate, true);
  html += comparisonRowHtml('効果率', players, p => p.efficiency, true);

  if (hasBlock) {
    html += comparisonSectionHtml('ブロック（全セット合計）', colspan);
    html += comparisonRowHtml('本数', players, p => p.block_attempts, false);
    html += comparisonRowHtml('得点', players, p => p.block_points, false);
    html += comparisonRowHtml('ワンチ', players, p => p.block_touch_own + p.block_touch_opp, false);
    html += comparisonRowHtml('失点', players, p => p.block_errors, false);
  }

  if (hasServe) {
    html += comparisonSectionHtml('サーブ（全セット合計）', colspan);
    html += comparisonRowHtml('打数', players, p => p.serve_attempts, false);
    html += comparisonRowHtml('エース', players, p => p.serve_aces, false);
    html += comparisonRowHtml('ミス', players, p => p.serve_errors, false);
    html += comparisonRowHtml('効果本数', players, p => p.serve_half_credit, false);
    html += comparisonRowHtml('エース率', players, p => p.serve_ace_rate, true);
    html += comparisonRowHtml('ミス率', players, p => p.serve_error_rate, true);
    html += comparisonRowHtml('効果率', players, p => p.serve_efficiency, true);
  }

  if (hasReceive) {
    html += comparisonSectionHtml('レシーブ（全セット合計）', colspan);
    html += comparisonRowHtml('本数', players, p => p.receive_attempts, false);
    html += comparisonRowHtml('Aパス', players, p => p.receive_a, false);
    html += comparisonRowHtml('Bパス', players, p => p.receive_b, false);
    html += comparisonRowHtml('Cパス', players, p => p.receive_c, false);
    html += comparisonRowHtml('Dパス', players, p => p.receive_d, false);
    html += comparisonRowHtml('ミス', players, p => p.receive_errors, false);
    html += comparisonRowHtml('返球率', players, p => p.receive_return_rate, true);
    html += comparisonRowHtml('A率', players, p => p.receive_a_rate, true);
    html += comparisonRowHtml('AB率', players, p => p.receive_ab_rate, true);
  }

  html += '</tbody></table></div>';
  return html;
}

function showComparison(nums) {
  const players = nums
    .map(n => DATA.total.find(pp => pp.number === n))
    .filter(Boolean);
  document.getElementById('comparisonCard').style.display = 'block';
  document.getElementById('comparisonTitle').textContent = `選手比較（${players.length}人）`;
  document.getElementById('comparisonBody').innerHTML = comparisonTableHtml(players);
}

function rotationTableHtml(rows) {
  let html = '<div style="overflow-x:auto"><table class="rotation-table" style="min-width:920px">';
  html += '<thead><tr><th>ローテ</th><th>打数</th><th>得点</th><th>ミス</th><th>決定率</th>';
  CATEGORY_ORDER.forEach(cat => {
    html += `<th>${CATEGORY_LABELS[cat]}本数</th><th>${CATEGORY_LABELS[cat]}決定率</th>`;
  });
  html += '</tr></thead><tbody>';
  rows.forEach(row => {
    const label = typeof row.rotation === 'number' ? `S${row.rotation}` : row.rotation;
    html += `<tr><td>${label}</td><td>${row.attempts}</td><td>${row.points}</td><td>${row.errors}</td><td>${pct(row.kill_rate)}</td>`;
    CATEGORY_ORDER.forEach(cat => {
      const c = row.categories[cat];
      html += `<td>${c.attempts}</td><td>${pct(c.kill_rate)}</td>`;
    });
    html += '</tr>';
  });
  html += '</tbody></table></div>';
  return html;
}

function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, ch => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[ch]));
}

function scoreChartSvg(points, title, ownLabel, opponentLabel) {
  const W = 640, H = 240;
  const padL = 32, padR = 34, padT = 14, padB = 14;
  const plotW = W - padL - padR;
  const plotH = H - padT - padB;
  const n = points.length;
  const maxScore = Math.max(1, ...points.map(p => Math.max(p.own, p.opponent)));
  const yMax = Math.ceil((maxScore + 1) / 5) * 5;
  const xAt = i => padL + (n <= 1 ? 0 : (i / (n - 1)) * plotW);
  const yAt = v => padT + plotH - (v / yMax) * plotH;

  let gridSvg = '';
  for (let t = 0; t <= yMax; t += 5) {
    const y = yAt(t);
    gridSvg += `<line x1="${padL}" y1="${y.toFixed(1)}" x2="${W - padR}" y2="${y.toFixed(1)}" `
             + `stroke="var(--grid)" stroke-width="1"/>`;
    gridSvg += `<text x="${padL - 8}" y="${(y + 4).toFixed(1)}" text-anchor="end" font-size="11" `
             + `fill="var(--text-muted)">${t}</text>`;
  }

  function polylineAndDots(key, colorVar) {
    const coords = points.map((p, i) => `${xAt(i).toFixed(1)},${yAt(p[key]).toFixed(1)}`).join(' ');
    const dots = points.map((p, i) => {
      const cx = xAt(i).toFixed(1), cy = yAt(p[key]).toFixed(1);
      return `<circle cx="${cx}" cy="${cy}" r="3" fill="${colorVar}" stroke="var(--surface-1)" `
           + `stroke-width="1"><title>${i}本目: ${p.own}-${p.opponent}</title></circle>`;
    }).join('');
    return `<polyline points="${coords}" fill="none" stroke="${colorVar}" stroke-width="2" `
         + `stroke-linejoin="round" stroke-linecap="round"/>${dots}`;
  }

  const last = points[points.length - 1];
  const lx = xAt(points.length - 1);
  let ownLy = yAt(last.own);
  let oppLy = yAt(last.opponent);
  // 2つの最終スコアが近いと数字が重なるので、近すぎる場合は少しずらす
  if (Math.abs(ownLy - oppLy) < 12) {
    if (ownLy <= oppLy) { ownLy -= 6; oppLy += 6; } else { ownLy += 6; oppLy -= 6; }
  }

  function endMark(cy, val, colorVar) {
    return `<circle cx="${lx.toFixed(1)}" cy="${yAt(val).toFixed(1)}" r="5" fill="${colorVar}" `
         + `stroke="var(--surface-1)" stroke-width="2"/>`
         + `<text x="${(lx + 9).toFixed(1)}" y="${(cy + 4).toFixed(1)}" font-size="12" font-weight="600" `
         + `fill="var(--text-primary)">${val}</text>`;
  }

  return `<div class="score-chart">`
       + `<div class="score-chart-title">${title}</div>`
       + `<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto;display:block">`
       + gridSvg
       + polylineAndDots('own', 'var(--series-blue)') + polylineAndDots('opponent', 'var(--series-orange)')
       + endMark(ownLy, last.own, 'var(--series-blue)') + endMark(oppLy, last.opponent, 'var(--series-orange)')
       + `</svg>`
       + `<div class="score-chart-legend">`
       + `<span class="legend-item"><span class="legend-swatch" style="background:var(--series-blue)"></span>`
       + `${escapeHtml(ownLabel)}</span>`
       + `<span class="legend-item"><span class="legend-swatch" style="background:var(--series-orange)"></span>`
       + `${escapeHtml(opponentLabel)}</span>`
       + `</div></div>`;
}

// ==================== セット開始時点のスタメン（P1〜P6） ====================
// 選手ランキングのすぐ上に、セットを選ぶボタンで切り替えて見せる（1試合を選んでいるときのみ）
function jerseySlotHtml(number, isServer) {
  const name = ROSTER[number] || `#${number}`;
  const badge = isServer ? '<span class="jersey-badge">S</span>' : '';
  return `<div class="jersey-slot">`
    + `<div class="jersey-name-box">${badge}${escapeHtml(name)}</div>`
    + `</div>`;
}

function startingLineupHtml(lineup) {
  if (!lineup || !lineup.length) return '<p class="hint">このセットのスタメン情報は取得できませんでした（この試合を追加した時点では未対応でした）。</p>';
  // コート図と同じ並びにする：前衛（ネット側）＝P4,P3,P2／後衛＝P5,P6,P1（サーブ位置）
  const front = [lineup[3], lineup[2], lineup[1]];
  const back = [lineup[4], lineup[5], lineup[0]];
  let html = '<div class="lineup-court">';
  html += '<div class="lineup-row">' + front.map(n => jerseySlotHtml(n, false)).join('') + '</div>';
  html += '<div class="lineup-row">' + back.map(n => jerseySlotHtml(n, n === lineup[0])).join('') + '</div>';
  html += '</div>';
  return html;
}

function selectStartingLineupSet(lineup, chipEl) {
  document.querySelectorAll('#startingLineupSetButtons .player-chip').forEach(c => c.classList.remove('selected'));
  chipEl.classList.add('selected');
  document.getElementById('startingLineupBody').innerHTML = startingLineupHtml(lineup);
}

function showStartingLineups(bySet) {
  const btnEl = document.getElementById('startingLineupSetButtons');
  btnEl.innerHTML = '';
  const withLineup = (bySet || []).filter(s => s.startingLineup && s.startingLineup.length);
  if (!withLineup.length) {
    document.getElementById('startingLineupBody').innerHTML =
      '<p class="hint">スタメン情報を取得できませんでした（この試合を追加した時点では未対応でした）。</p>';
    return;
  }
  withLineup.forEach((s, idx) => {
    const chip = document.createElement('button');
    chip.className = 'player-chip';
    chip.textContent = `第${s.setNumber}セット`;
    chip.addEventListener('click', () => selectStartingLineupSet(s.startingLineup, chip));
    if (idx === 0) chip.classList.add('selected');
    btnEl.appendChild(chip);
  });
  document.getElementById('startingLineupBody').innerHTML = startingLineupHtml(withLineup[0].startingLineup);
}

function selectScoreProgressionSet(s, chipEl) {
  document.querySelectorAll('#scoreProgressionSetButtons .player-chip').forEach(c => c.classList.remove('selected'));
  chipEl.classList.add('selected');
  document.getElementById('scoreProgressionChartArea').innerHTML =
    scoreChartSvg(s.points, `第${s.setNumber}セット`, DATA.teamLabel, DATA.opponent);
}

function showScoreProgressionCharts() {
  const btnEl = document.getElementById('scoreProgressionSetButtons');
  const chartEl = document.getElementById('scoreProgressionChartArea');
  btnEl.innerHTML = '';
  DATA.scoreProgression.bySet.forEach((s, idx) => {
    const chip = document.createElement('button');
    chip.className = 'player-chip';
    chip.textContent = `第${s.setNumber}セット`;
    chip.addEventListener('click', () => selectScoreProgressionSet(s, chip));
    if (idx === 0) chip.classList.add('selected');
    btnEl.appendChild(chip);
  });
  if (DATA.scoreProgression.bySet.length) {
    const first = DATA.scoreProgression.bySet[0];
    chartEl.innerHTML = scoreChartSvg(first.points, `第${first.setNumber}セット`, DATA.teamLabel, DATA.opponent);
  } else {
    chartEl.innerHTML = '';
  }
}

// 試合を2つ以上選んでいるとき：選んだ全試合・全セットをボタンで並べて、どれでも選べるようにする
// （以前は最初に選んだ試合しか出せなかったのを2026-08-29に修正）
function selectScoreProgressionCombinedItem(item, chipEl) {
  document.querySelectorAll('#scoreProgressionSetButtons .player-chip').forEach(c => c.classList.remove('selected'));
  chipEl.classList.add('selected');
  document.getElementById('scoreProgressionChartArea').innerHTML =
    scoreChartSvg(item.points, item.setLabel, item.teamLabel, item.opponent);
}

function showScoreProgressionChartsCombined(entries, dataList) {
  const btnEl = document.getElementById('scoreProgressionSetButtons');
  const chartEl = document.getElementById('scoreProgressionChartArea');
  btnEl.innerHTML = '';
  const items = [];
  dataList.forEach(d => {
    (d.scoreProgression.bySet || []).forEach(s => {
      items.push({
        label: `${d.opponent}（${formatMatchDate(d.matchDate)}）第${s.setNumber}セット`,
        points: s.points,
        setLabel: `${d.opponent} 第${s.setNumber}セット`,
        teamLabel: d.teamLabel,
        opponent: d.opponent,
      });
    });
  });
  items.forEach((item, idx) => {
    const chip = document.createElement('button');
    chip.className = 'player-chip';
    chip.textContent = item.label;
    chip.addEventListener('click', () => selectScoreProgressionCombinedItem(item, chip));
    if (idx === 0) chip.classList.add('selected');
    btnEl.appendChild(chip);
  });
  if (items.length) {
    chartEl.innerHTML = scoreChartSvg(items[0].points, items[0].setLabel, items[0].teamLabel, items[0].opponent);
  } else {
    chartEl.innerHTML = '';
  }
}

// ==================== 得点パターン：試合ごと／全体の切り替え ====================
function selectScoringPatternView(bySet, chipEl) {
  document.querySelectorAll('#scoringPatternMatchButtons .player-chip').forEach(c => c.classList.remove('selected'));
  chipEl.classList.add('selected');
  document.getElementById('scoringPatternBody').innerHTML = scoringPatternHtml(computeScoringPattern(bySet));
}

function showScoringPattern() {
  document.getElementById('scoringPatternMatchButtons').innerHTML = '';
  document.getElementById('scoringPatternBody').innerHTML =
    scoringPatternHtml(computeScoringPattern(DATA.scoreProgression.bySet));
}

function showScoringPatternCombined(entries, dataList) {
  const btnEl = document.getElementById('scoringPatternMatchButtons');
  btnEl.innerHTML = '';
  const items = dataList.map(d => ({
    label: `${d.opponent}（${formatMatchDate(d.matchDate)}）`,
    bySet: d.scoreProgression.bySet,
  }));
  items.push({ label: '全体', bySet: dataList.flatMap(d => d.scoreProgression.bySet) });
  items.forEach((item, idx) => {
    const chip = document.createElement('button');
    chip.className = 'player-chip';
    chip.textContent = item.label;
    chip.addEventListener('click', () => selectScoringPatternView(item.bySet, chip));
    if (idx === items.length - 1) chip.classList.add('selected');
    btnEl.appendChild(chip);
  });
  const last = items[items.length - 1];
  document.getElementById('scoringPatternBody').innerHTML = scoringPatternHtml(computeScoringPattern(last.bySet));
}

// ==================== セットごとの流れハイライト：試合ごと／全体の切り替え ====================
function selectRunHighlightsView(scoreProgression, chipEl) {
  document.querySelectorAll('#runHighlightsMatchButtons .player-chip').forEach(c => c.classList.remove('selected'));
  chipEl.classList.add('selected');
  document.getElementById('runHighlightsBody').innerHTML = runHighlightsHtml(scoreProgression);
}

function showRunHighlights() {
  document.getElementById('runHighlightsMatchButtons').innerHTML = '';
  document.getElementById('runHighlightsBody').innerHTML = runHighlightsHtml(DATA.scoreProgression);
}

function showRunHighlightsCombined(entries, dataList) {
  const btnEl = document.getElementById('runHighlightsMatchButtons');
  btnEl.innerHTML = '';
  const items = dataList.map(d => ({
    label: `${d.opponent}（${formatMatchDate(d.matchDate)}）`,
    sp: d.scoreProgression,
  }));
  items.push({ label: '全体', sp: { bySet: dataList.flatMap(d => d.scoreProgression.bySet) } });
  items.forEach((item, idx) => {
    const chip = document.createElement('button');
    chip.className = 'player-chip';
    chip.textContent = item.label;
    chip.addEventListener('click', () => selectRunHighlightsView(item.sp, chip));
    if (idx === items.length - 1) chip.classList.add('selected');
    btnEl.appendChild(chip);
  });
  const last = items[items.length - 1];
  document.getElementById('runHighlightsBody').innerHTML = runHighlightsHtml(last.sp);
}

// ==================== セットごとの流れハイライト（連続得点・点差） ====================
// ラン(連続得点/連続失点)の最大記録を数えつつ、その場面に絡んだ自チームの
// 選手番号も一緒に記録する（得点に直接絡んだ選手／自チームのミスで失点に
// 絡んだ選手）。相手のミスによる得点など、選手を特定できない場面は含めない。
function computeRunHighlights(scoreProgression) {
  const perSet = (scoreProgression.bySet || []).map(s => {
    const pts = s.points || [];
    let lastOwn = 0, lastOpp = 0;
    let curRunTeam = null, curRunLen = 0, curRunPlayers = [];
    let bestOwnRun = 0, bestOppRun = 0;
    let bestOwnRunPlayers = [], bestOppRunPlayers = [];
    pts.forEach((p, i) => {
      if (i === 0) { lastOwn = p.own; lastOpp = p.opponent; return; }
      const ownGained = p.own - lastOwn;
      const oppGained = p.opponent - lastOpp;
      if (ownGained > 0) {
        if (curRunTeam === 'own') { curRunLen += 1; } else { curRunTeam = 'own'; curRunLen = 1; curRunPlayers = []; }
        if (p.scorerNumber !== null && p.scorerNumber !== undefined) curRunPlayers.push(p.scorerNumber);
        if (curRunLen > bestOwnRun) { bestOwnRun = curRunLen; bestOwnRunPlayers = curRunPlayers.slice(); }
      } else if (oppGained > 0) {
        if (curRunTeam === 'opp') { curRunLen += 1; } else { curRunTeam = 'opp'; curRunLen = 1; curRunPlayers = []; }
        if (p.scorerNumber !== null && p.scorerNumber !== undefined) curRunPlayers.push(p.scorerNumber);
        if (curRunLen > bestOppRun) { bestOppRun = curRunLen; bestOppRunPlayers = curRunPlayers.slice(); }
      }
      lastOwn = p.own; lastOpp = p.opponent;
    });
    const finalOwn = pts.length ? pts[pts.length - 1].own : 0;
    const finalOpp = pts.length ? pts[pts.length - 1].opponent : 0;
    return {
      setNumber: s.setNumber,
      finalOwn, finalOpp,
      margin: Math.abs(finalOwn - finalOpp),
      won: finalOwn > finalOpp,
      bestOwnRun, bestOppRun,
      bestOwnRunPlayers, bestOppRunPlayers,
    };
  });
  let overallBestOwn = { setNumber: null, run: 0, players: [] };
  let overallBestOpp = { setNumber: null, run: 0, players: [] };
  perSet.forEach(s => {
    if (s.bestOwnRun > overallBestOwn.run) overallBestOwn = { setNumber: s.setNumber, run: s.bestOwnRun, players: s.bestOwnRunPlayers };
    if (s.bestOppRun > overallBestOpp.run) overallBestOpp = { setNumber: s.setNumber, run: s.bestOppRun, players: s.bestOppRunPlayers };
  });
  return { perSet, overallBestOwn, overallBestOpp };
}

// 選手番号のリストを「菊田(2)、小林」のような表示用文字列にする（登場回数の多い順）
function formatRunPlayers(numbers) {
  if (!numbers || !numbers.length) return '';
  const counts = {};
  numbers.forEach(n => { counts[n] = (counts[n] || 0) + 1; });
  return Object.keys(counts).map(Number)
    .sort((a, b) => counts[b] - counts[a])
    .map(n => `${ROSTER[n] || '#' + n}${counts[n] > 1 ? `(${counts[n]})` : ''}`)
    .join('、');
}

function runHighlightsHtml(scoreProgression) {
  const { perSet, overallBestOwn, overallBestOpp } = computeRunHighlights(scoreProgression);
  if (!perSet.length) return '<p class="hint">データがありません。</p>';
  const ownPlayersText = formatRunPlayers(overallBestOwn.players);
  const oppPlayersText = formatRunPlayers(overallBestOpp.players);
  let html = '<div class="stat-tiles" style="margin-bottom:6px">';
  html += `<div class="stat-tile"><div class="k">自チーム最大連続得点</div><div class="v">${overallBestOwn.run}点</div></div>`;
  html += `<div class="stat-tile"><div class="k">相手最大連続得点</div><div class="v">${overallBestOpp.run}点</div></div>`;
  html += '</div>';
  if (ownPlayersText || oppPlayersText) {
    html += '<p class="hint" style="margin-top:0;margin-bottom:14px">';
    if (ownPlayersText) html += `自チーム最大連続得点で決めた選手：${escapeHtml(ownPlayersText)}`;
    if (ownPlayersText && oppPlayersText) html += '　／　';
    if (oppPlayersText) html += `相手最大連続得点で自チームのミスが絡んだ選手：${escapeHtml(oppPlayersText)}`;
    html += '</p>';
  }
  html += '<div style="overflow-x:auto"><table class="rotation-table"><thead><tr>'
    + '<th>セット</th><th>最終スコア</th><th>点差</th><th>結果</th><th>自チーム最大連続</th><th>相手最大連続</th>'
    + '</tr></thead><tbody>';
  perSet.forEach(s => {
    const ownP = formatRunPlayers(s.bestOwnRunPlayers);
    const oppP = formatRunPlayers(s.bestOppRunPlayers);
    html += `<tr><td>第${s.setNumber}セット</td>`
      + `<td>${s.finalOwn} - ${s.finalOpp}</td>`
      + `<td>${s.margin}点差</td>`
      + `<td>${s.won ? '○' : '×'}</td>`
      + `<td>${s.bestOwnRun}点${ownP ? `<br><span class="hint" style="font-size:11px">${escapeHtml(ownP)}</span>` : ''}</td>`
      + `<td>${s.bestOppRun}点${oppP ? `<br><span class="hint" style="font-size:11px">${escapeHtml(oppP)}</span>` : ''}</td></tr>`;
  });
  html += '</tbody></table></div>';
  return html;
}

// トス配分（％）：レフト／ライト／ミドル／バックアタック／その他に、打数のうち何%が上がったかを表示する
// （rowsの中の「合計」行＝その選択範囲全体の内訳を使う。2026-08-29追加）
function tossDistributionHtml(totalRow) {
  if (!totalRow || !totalRow.attempts) return '<p class="hint">データがありません。</p>';
  let html = '<div class="stat-tiles">';
  CATEGORY_ORDER.forEach(cat => {
    const c = totalRow.categories[cat];
    const share = totalRow.attempts ? c.attempts / totalRow.attempts : null;
    html += `<div class="stat-tile"><div class="k">${CATEGORY_LABELS[cat]}</div><div class="v">${pctCount(share, c.attempts, totalRow.attempts)}</div></div>`;
  });
  html += '</div>';
  return html;
}

// 打法（％）：強打／フェイントに、打数のうち何%が該当するかを表示する
// （スパイクのタイプ文字H=強打／T=フェイントより。2026-08-29追加）
function attackTechniqueHtml(totalRow) {
  if (!totalRow || !totalRow.attempts) return '<p class="hint">データがありません。</p>';
  let html = '<div class="stat-tiles">';
  TECH_ORDER.forEach(tech => {
    const t = totalRow.techniques[tech];
    const share = totalRow.attempts ? t.attempts / totalRow.attempts : null;
    html += `<div class="stat-tile"><div class="k">${TECH_LABELS[tech]}</div><div class="v">${pctCount(share, t.attempts, totalRow.attempts)}</div></div>`;
  });
  html += '</div>';
  return html;
}

// 相手が決めた攻撃（アタックで得点したものだけ）のコース内訳（レフト/ミドル/ライト/バックアタック/その他）
// トス配分・打法別と違い、分母は「打数」ではなく「相手が決めた本数」（2026-09-03追加）。
function opponentCourseHtml(totalRow) {
  if (!totalRow || !totalRow.points) return '<p class="hint">相手の決定データがありません。</p>';
  let html = '<div class="stat-tiles">';
  CATEGORY_ORDER.forEach(cat => {
    const c = totalRow.categories[cat];
    const share = totalRow.points ? c.points / totalRow.points : null;
    html += `<div class="stat-tile"><div class="k">${CATEGORY_LABELS[cat]}</div><div class="v">${pctCount(share, c.points, totalRow.points)}</div></div>`;
  });
  html += '</div>';
  return html;
}

// 相手が決めたコースをコート図（矢印）で見せる簡易版（2026-09-04追加）。
// 1本ごとの正確な着地位置ではなく、レフト／ミドル／ライト／バックアタックそれぞれの
// 「代表的な位置」に矢印を引き、本数・割合をラベルとして添える。
// 「全体」＝4方向まとめて表示、コース名を選ぶとそのコースだけを表示する。
const OPPONENT_COURSE_FILTER_OPTIONS = [
  ['all', '全体'], ['left', 'レフト'], ['middle', 'ミドル'], ['right', 'ライト'], ['pipe', 'バックアタック'],
];
const OPPONENT_COURSE_POINTS = {
  left: { x: 45, y: 175 },
  middle: { x: 100, y: 245 },
  right: { x: 155, y: 175 },
  pipe: { x: 100, y: 120 },
};

// こうせいさんがVolleyStationの「決定率100%（＝決めた攻撃だけ）」の画面をスクリーンショットで送ってくれたものから、
// 黒い点（実際の着地位置）の座標を読み取って登録した実データ（2026-09-04追加）。
// あくまで画像から読み取った近似値で、1本1本の正確な数値データではない点に注意。
// キー：dvBuildMatchSlug(matchDate, opponent) と同じ形式（例: '2026-08-30_千葉ベルズ'）。
// カテゴリごとの本数は、画面外のコンビも含む正式な集計（stat-tilesの数字）と一致しない場合がある
// （スクリーンショットに写っていたコンビだけを読み取っているため）。
const OPPONENT_COURSE_REAL_POINTS = {
  '2026-08-30_千葉ベルズ': {
    left: [[62.7,289.3],[81.8,286],[119.2,247.8],[99.3,246.4],[50.4,240.8],[167.5,230.6],[61.2,225.2],[66.8,210.2],[136.2,201.8],[46.6,193.3],[69.6,190.7],[126.8,189.9],[20.7,151.8],[141.6,151.6],[56.8,136.2],[111.2,147.3],[169.2,125.9],[114.4,285.3],[170.1,281.2],[182.6,250.9],[86.7,247.2],[136.4,232.8],[78.7,206]],
    middle: [[162.1,202.1],[149.6,201.5],[120.4,195.5],[145,190.7],[76.2,158.1],[101.4,119],[169,236.7],[71.3,208.5],[142.3,188.1],[81.6,236.1],[146.4,212.3],[89.4,201.7],[105.1,199.7],[119.7,181.1],[35.4,179.9]],
    right: [[65.8,278.4],[56.8,247.6],[113.2,243],[70.1,227.3],[133.1,187.9],[79.8,188.2],[40,112.6]],
  },
};

let CURRENT_OPP_COURSE_ROW = null;
let CURRENT_OPP_COURSE_FILTER = 'all';
let CURRENT_OPP_COURSE_REAL_KEY = null;

function opponentCourseDiagramSvg(totalRow, filterCat, realKey) {
  if (!totalRow || !totalRow.points) return '';
  const catsToShow = filterCat === 'all' ? ['left', 'right', 'middle', 'pipe'] : [filterCat];
  const originX = 100, originY = 70;
  const realData = realKey ? OPPONENT_COURSE_REAL_POINTS[realKey] : null;
  let arrows = '';
  let realNote = '';
  const realLegendParts = [];
  catsToShow.forEach(cat => {
    const c = totalRow.categories[cat];
    if (!c) return;
    const realPts = realData ? realData[cat] : null;
    if (realPts && realPts.length) {
      // 実データ（スクリーンショットから読み取った実際の着地位置）がある場合は、1本ずつ点を打つ
      realPts.forEach(([px, py]) => {
        arrows += `
          <line x1="${originX}" y1="${originY}" x2="${px}" y2="${py}" stroke="var(--series-blue)" stroke-width="1" stroke-dasharray="2,2" opacity="0.55" />
          <circle cx="${px}" cy="${py}" r="3" fill="var(--series-blue)" opacity="0.85" />`;
      });
      realLegendParts.push(`${CATEGORY_LABELS[cat]}（実データ${realPts.length}本）`);
      realNote = '<p class="hint" style="margin-top:6px">実データがある方向は、スクリーンショットから読み取った実際の着地位置を1本ずつ表示しています（読み取りのため近似値です）。それ以外は代表点での簡易表示です。</p>';
    } else {
      const pt = OPPONENT_COURSE_POINTS[cat];
      if (!pt) return;
      const pct = totalRow.points ? Math.round((c.points / totalRow.points) * 1000) / 10 : 0;
      arrows += `
        <line x1="${originX}" y1="${originY}" x2="${pt.x}" y2="${pt.y}" stroke="var(--series-blue)" stroke-width="1.6" stroke-dasharray="3,3" />
        <circle cx="${pt.x}" cy="${pt.y}" r="5" fill="var(--series-blue)" />
        <text x="${pt.x}" y="${pt.y - 11}" text-anchor="middle" font-size="11" font-weight="700" fill="var(--text-primary)">${CATEGORY_LABELS[cat]}</text>
        <text x="${pt.x}" y="${pt.y + 19}" text-anchor="middle" font-size="10" fill="var(--text-secondary)">${c.points}本（${pct}%）</text>`;
    }
  });
  const legendHtml = realLegendParts.length
    ? `<p style="text-align:center;font-size:11px;font-weight:700;color:var(--text-primary);margin:6px 0 0">${realLegendParts.join('　')}</p>`
    : '';
  return `
    <svg viewBox="0 0 200 300" style="width:100%;max-width:260px;display:block;margin:0 auto" xmlns="http://www.w3.org/2000/svg">
      <rect x="10" y="20" width="180" height="270" fill="none" stroke="var(--border)" stroke-width="1.5" />
      <line x1="70" y1="70" x2="70" y2="290" stroke="var(--grid)" stroke-width="1" stroke-dasharray="2,3" />
      <line x1="130" y1="70" x2="130" y2="290" stroke="var(--grid)" stroke-width="1" stroke-dasharray="2,3" />
      <line x1="10" y1="143" x2="190" y2="143" stroke="var(--grid)" stroke-width="1" />
      <line x1="10" y1="217" x2="190" y2="217" stroke="var(--grid)" stroke-width="1" />
      <line x1="10" y1="70" x2="190" y2="70" stroke="var(--text-primary)" stroke-width="2.5" />
      <circle cx="${originX}" cy="${originY}" r="4" fill="#e5484d" />
      ${arrows}
    </svg>${legendHtml}${realNote}`;
}

function setOpponentCourseFilter(key) {
  CURRENT_OPP_COURSE_FILTER = key;
  renderOpponentCourseSection(CURRENT_OPP_COURSE_ROW, CURRENT_OPP_COURSE_REAL_KEY);
}

function renderOpponentCourseSection(totalRow, realKey) {
  CURRENT_OPP_COURSE_ROW = totalRow;
  CURRENT_OPP_COURSE_REAL_KEY = realKey || null;
  const filterBtnsEl = document.getElementById('opponentCourseFilterButtons');
  if (filterBtnsEl) {
    filterBtnsEl.innerHTML = '';
    OPPONENT_COURSE_FILTER_OPTIONS.forEach(([key, label]) => {
      const chip = document.createElement('button');
      chip.type = 'button';
      chip.className = 'player-chip' + (CURRENT_OPP_COURSE_FILTER === key ? ' selected' : '');
      chip.textContent = label;
      chip.addEventListener('click', () => setOpponentCourseFilter(key));
      filterBtnsEl.appendChild(chip);
    });
  }
  const oppCourseBody = document.getElementById('opponentCourseBody');
  if (oppCourseBody) oppCourseBody.innerHTML = opponentCourseHtml(totalRow);
  const diagramEl = document.getElementById('opponentCourseDiagram');
  if (diagramEl) diagramEl.innerHTML = opponentCourseDiagramSvg(totalRow, CURRENT_OPP_COURSE_FILTER, CURRENT_OPP_COURSE_REAL_KEY);
}

function selectRotationSet(rows, chipEl) {
  document.querySelectorAll('#rotationSetButtons .player-chip').forEach(c => c.classList.remove('selected'));
  chipEl.classList.add('selected');
  document.getElementById('rotationTables').innerHTML = rotationTableHtml(rows);
  const totalRow = rows.find(r => r.rotation === '合計') || rows[rows.length - 1];
  document.getElementById('tossDistributionBody').innerHTML = tossDistributionHtml(totalRow);
  const techBody = document.getElementById('attackTechniqueBody');
  if (techBody) techBody.innerHTML = attackTechniqueHtml(totalRow);
}

function showRotationTables() {
  document.getElementById('rotationCardLabel').textContent = 'ローテーション別 攻撃タイプ分布（チーム全体）';
  document.getElementById('rotationCardHint').textContent =
    'セット開始時のローテーションをS1とし、自チームが1回転するたびにS2→S3…と数えています。攻撃タイプ（レフト／ライト／ミドル／バックアタック）は使用した攻撃コンビの分類によるものです。セットを選んでください。';
  const btnEl = document.getElementById('rotationSetButtons');
  btnEl.innerHTML = '';
  const items = DATA.rotation.bySet.map((s, i) => ({
    label: `第${s.setNumber}セット`,
    rows: s.rows,
  }));
  items.push({
    label: '全セット合計',
    rows: DATA.rotation.total.rows,
  });
  items.forEach((item, idx) => {
    const chip = document.createElement('button');
    chip.className = 'player-chip';
    chip.textContent = item.label;
    chip.addEventListener('click', () => selectRotationSet(item.rows, chip));
    if (idx === items.length - 1) chip.classList.add('selected');
    btnEl.appendChild(chip);
  });
  const last = items[items.length - 1];
  selectRotationSet(last.rows, btnEl.children[items.length - 1]);
}

// 「相手が決めた攻撃のコース」カード専用のセット切り替え（rotationCardとは別カードになったので、
// セット選択ボタンも別に持つ。2026-09-04追加）
function selectOpponentCourseSet(rows, chipEl, realKey) {
  document.querySelectorAll('#opponentCourseSetButtons .player-chip').forEach(c => c.classList.remove('selected'));
  if (chipEl) chipEl.classList.add('selected');
  const totalRow = (rows && rows.length) ? (rows.find(r => r.rotation === '合計') || rows[rows.length - 1]) : null;
  renderOpponentCourseSection(totalRow, realKey);
}

function showOpponentCourseSets() {
  const btnEl = document.getElementById('opponentCourseSetButtons');
  if (!btnEl) return;
  btnEl.innerHTML = '';
  const matchSlug = dvBuildMatchSlug(DATA.matchDate, DATA.opponent);
  const items = (DATA.opponentRotation ? DATA.opponentRotation.bySet : []).map(s => ({
    label: `第${s.setNumber}セット`,
    rows: s.rows,
    realKey: null, // 実データはセット別ではなく試合合計でしか持っていない
  }));
  items.push({
    label: '全セット合計',
    rows: DATA.opponentRotation ? DATA.opponentRotation.total.rows : [],
    realKey: matchSlug,
  });
  items.forEach((item, idx) => {
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = 'player-chip';
    chip.textContent = item.label;
    chip.addEventListener('click', () => selectOpponentCourseSet(item.rows, chip, item.realKey));
    if (idx === items.length - 1) chip.classList.add('selected');
    btnEl.appendChild(chip);
  });
  const last = items[items.length - 1];
  selectOpponentCourseSet(last.rows, btnEl.children[items.length - 1], last.realKey);
}

// 試合を2つ以上選んでいるとき用（rotationCardのshowRotationTablesCombinedと同じパターン）
function showOpponentCourseSetsCombined(entries, dataList) {
  const btnEl = document.getElementById('opponentCourseSetButtons');
  if (!btnEl) return;
  btnEl.innerHTML = '';
  const items = dataList.map(d => ({
    label: `${d.opponent}（${formatMatchDate(d.matchDate)}）`,
    rows: d.opponentRotation ? d.opponentRotation.total.rows : [],
    realKey: dvBuildMatchSlug(d.matchDate, d.opponent), // その試合単体を見ているときだけ実データを使う
  }));
  items.push({
    label: '全体',
    rows: combineRotationRows(dataList.map(d => (d.opponentRotation && d.opponentRotation.total.rows) || [])),
    realKey: null, // 複数試合の合算では実データ（1試合分）は使わない
  });
  items.forEach((item, idx) => {
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = 'player-chip';
    chip.textContent = item.label;
    chip.addEventListener('click', () => selectOpponentCourseSet(item.rows, chip, item.realKey));
    if (idx === items.length - 1) chip.classList.add('selected');
    btnEl.appendChild(chip);
  });
  const last = items[items.length - 1];
  selectOpponentCourseSet(last.rows, btnEl.children[items.length - 1], last.realKey);
}

// 試合を2つ以上選んでいるとき：試合ごとのボタン＋「全体」（選んだ試合すべての合計）を並べて選べるようにする
// （以前は選んだ試合すべての合計しか出せなかったのを2026-08-29に修正）
function showRotationTablesCombined(entries, dataList) {
  document.getElementById('rotationCardLabel').textContent =
    `ローテーション別 攻撃タイプ分布（チーム全体・選んだ${entries.length}試合）`;
  document.getElementById('rotationCardHint').textContent =
    'セット開始時のローテーションをS1とし、自チームが1回転するたびにS2→S3…と数えています。攻撃タイプ（レフト／ライト／ミドル／バックアタック）は使用した攻撃コンビの分類によるものです。試合ごと、または選んだ試合すべての合計（全体）を選べます。';
  const btnEl = document.getElementById('rotationSetButtons');
  btnEl.innerHTML = '';
  const items = dataList.map(d => ({
    label: `${d.opponent}（${formatMatchDate(d.matchDate)}）`,
    rows: d.rotation.total.rows,
  }));
  items.push({
    label: '全体',
    rows: combineRotationRows(dataList.map(d => d.rotation.total.rows)),
  });
  items.forEach((item, idx) => {
    const chip = document.createElement('button');
    chip.className = 'player-chip';
    chip.textContent = item.label;
    chip.addEventListener('click', () => selectRotationSet(item.rows, chip));
    if (idx === items.length - 1) chip.classList.add('selected');
    btnEl.appendChild(chip);
  });
  const last = items[items.length - 1];
  selectRotationSet(last.rows, btnEl.children[items.length - 1]);
}

function selectTeam(chipEl) {
  selectedPlayerNumbers.clear();
  document.querySelectorAll('#playerChecklist input[type="checkbox"]').forEach(cb => { cb.checked = false; });
  updatePlayerListToggleLabel();
  clearDropZoneStatus();
  markSelected(chipEl);
  document.getElementById('comparisonCard').style.display = 'none';
  showDetail(DATA.team.name, 'チーム合計 ／ 全セット合計', DATA.team.total, DATA.team.sets);
  refreshMatchComparisonIfNeeded();
}

function selectSideOutBreakSet(s, chipEl) {
  document.querySelectorAll('#sideOutBreakSetButtons .player-chip').forEach(c => c.classList.remove('selected'));
  chipEl.classList.add('selected');
  document.getElementById('sideOutBreakTiles').innerHTML =
      `<div class="stat-tile"><div class="k">サーブ本数</div><div class="v">${s.serve_total}</div></div>`
    + `<div class="stat-tile"><div class="k">ブレイク率</div><div class="v">${pct(s.break_rate)}</div></div>`
    + `<div class="stat-tile"><div class="k">レシーブ本数</div><div class="v">${s.receive_total}</div></div>`
    + `<div class="stat-tile"><div class="k">サイドアウト率</div><div class="v">${pct(s.side_out_rate)}</div></div>`;
}

function showSideOutBreakTable() {
  document.getElementById('sideOutBreakLabel').textContent = 'サイドアウト率・ブレイク率（チーム全体）';
  document.getElementById('sideOutBreakHint').textContent =
    'サイドアウト率＝相手のサーブを受けたラリーで勝った割合／ブレイク率＝自分たちのサーブのラリーで勝った割合です。セットを選んでください。';
  document.getElementById('sideOutBreakByRotationLabel').textContent = 'ローテーション別（全セット合計）';
  const btnEl = document.getElementById('sideOutBreakSetButtons');
  btnEl.innerHTML = '';
  const items = DATA.sideOutBreak.bySet.map(s => ({ label: `第${s.setNumber}セット`, data: s }));
  items.push({ label: '合計', data: DATA.sideOutBreak.total });
  items.forEach((item, idx) => {
    const chip = document.createElement('button');
    chip.className = 'player-chip';
    chip.textContent = item.label;
    chip.addEventListener('click', () => selectSideOutBreakSet(item.data, chip));
    if (idx === items.length - 1) chip.classList.add('selected');
    btnEl.appendChild(chip);
  });
  const last = items[items.length - 1];
  selectSideOutBreakSet(last.data, btnEl.children[items.length - 1]);
  document.getElementById('sideOutBreakByRotationBody').innerHTML = sideOutBreakByRotationHtml(DATA.sideOutBreak.byRotation || []);
  const insight = rotationInsightText(DATA.sideOutBreak.byRotation || []);
  document.getElementById('rotationInsightLine').style.display = insight ? '' : 'none';
  document.getElementById('rotationInsightLine').textContent = insight;
}

// 試合を2つ以上選んでいるとき：試合ごとのボタン＋「全体」（選んだ試合すべての合計）を並べて選べるようにする
// （以前は選んだ試合すべての合計しか出せなかったのを2026-08-29に修正）
function selectSideOutBreakView(item, chipEl) {
  document.querySelectorAll('#sideOutBreakSetButtons .player-chip').forEach(c => c.classList.remove('selected'));
  chipEl.classList.add('selected');
  document.getElementById('sideOutBreakTiles').innerHTML =
      `<div class="stat-tile"><div class="k">サーブ本数</div><div class="v">${item.data.serve_total}</div></div>`
    + `<div class="stat-tile"><div class="k">ブレイク率</div><div class="v">${pct(item.data.break_rate)}</div></div>`
    + `<div class="stat-tile"><div class="k">レシーブ本数</div><div class="v">${item.data.receive_total}</div></div>`
    + `<div class="stat-tile"><div class="k">サイドアウト率</div><div class="v">${pct(item.data.side_out_rate)}</div></div>`;
  document.getElementById('sideOutBreakByRotationBody').innerHTML = sideOutBreakByRotationHtml(item.byRotation || []);
  const insight = rotationInsightText(item.byRotation || []);
  document.getElementById('rotationInsightLine').style.display = insight ? '' : 'none';
  document.getElementById('rotationInsightLine').textContent = insight;
}

function showSideOutBreakTableCombined(entries, dataList) {
  document.getElementById('sideOutBreakLabel').textContent =
    `サイドアウト率・ブレイク率（チーム全体・選んだ${entries.length}試合）`;
  document.getElementById('sideOutBreakHint').textContent =
    'サイドアウト率＝相手のサーブを受けたラリーで勝った割合／ブレイク率＝自分たちのサーブのラリーで勝った割合です。試合ごと、または選んだ試合すべての合計（全体）を選べます。';
  document.getElementById('sideOutBreakByRotationLabel').textContent = `ローテーション別（選んだ${entries.length}試合）`;
  const btnEl = document.getElementById('sideOutBreakSetButtons');
  btnEl.innerHTML = '';
  const items = dataList.map(d => ({
    label: `${d.opponent}（${formatMatchDate(d.matchDate)}）`,
    data: d.sideOutBreak.total,
    byRotation: d.sideOutBreak.byRotation || [],
  }));
  items.push({
    label: '全体',
    data: combineSideOutBreak(dataList.map(d => d.sideOutBreak.total)),
    byRotation: combineSideOutBreakByRotation(dataList.map(d => d.sideOutBreak.byRotation || [])),
  });
  items.forEach((item, idx) => {
    const chip = document.createElement('button');
    chip.className = 'player-chip';
    chip.textContent = item.label;
    chip.addEventListener('click', () => selectSideOutBreakView(item, chip));
    if (idx === items.length - 1) chip.classList.add('selected');
    btnEl.appendChild(chip);
  });
  const last = items[items.length - 1];
  selectSideOutBreakView(last, btnEl.children[items.length - 1]);
}

// ==================== プレー映像（統計を押すとその場面の動画がすぐ再生される機能） ====================
// YouTubeの動画URL（通常URL・短縮URLどちらも）からvideo IDだけを取り出す。
// YouTube以外や、形式が読み取れないURLの場合はnullを返す。
function extractYouTubeId(url) {
  try {
    const u = new URL(url);
    if (u.hostname.includes('youtu.be')) {
      return u.pathname.replace(/^\//, '') || null;
    }
    if (u.hostname.includes('youtube.com')) {
      if (u.searchParams.get('v')) return u.searchParams.get('v');
      const m = u.pathname.match(/\/embed\/([^/?]+)/);
      if (m) return m[1];
    }
  } catch (e) {
    // 何もしない（下でnullを返す）
  }
  return null;
}

// 指定した秒数から、そのプレーの動画をすぐに再生する。
// YouTubeのURLなら埋め込みプレーヤーで自動再生、それ以外は新しいタブで開く。
function playVideoClip(seconds, btnEl) {
  const video = DATA && DATA.video;
  if (!video || !video.url) return;
  const videoId = extractYouTubeId(video.url);
  const wrap = document.getElementById('videoPlayerWrap');
  const frame = document.getElementById('videoPlayerFrame');
  if (videoId && wrap && frame) {
    frame.src = `https://www.youtube.com/embed/${videoId}?start=${seconds}&autoplay=1`;
    wrap.style.display = '';
    wrap.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  } else {
    window.open(videoLinkForSeconds(video.url, seconds), '_blank', 'noopener');
  }
  document.querySelectorAll('.video-clip-link.playing').forEach(el => el.classList.remove('playing'));
  if (btnEl) btnEl.classList.add('playing');
}

// YouTube以外のホスティング（埋め込み再生に対応できない場合）向けのフォールバック。
// URLに「t=◯◯s」を付け足して、その秒数から再生が始まるリンクを作る。
function videoLinkForSeconds(url, seconds) {
  try {
    const u = new URL(url);
    u.searchParams.set('t', seconds + 's');
    return u.toString();
  } catch (e) {
    return url;
  }
}

const VIDEO_CLIP_CATEGORIES = [
  { key: 'attack_kill', label: 'スパイク決定', skill: 'A', evaluation: '#' },
  { key: 'attack_error', label: 'スパイクミス', skill: 'A', evaluation: '=' },
  { key: 'serve_ace', label: 'サーブエース', skill: 'S', evaluation: '#' },
  { key: 'serve_error', label: 'サーブミス', skill: 'S', evaluation: '=' },
  { key: 'receive_a', label: 'レシーブ（Aパス）', skill: 'R', evaluation: '#' },
  { key: 'receive_error', label: 'レシーブミス', skill: 'R', evaluation: '=' },
  { key: 'block_point', label: 'ブロック得点', skill: 'B', evaluation: '#' },
];

// 選手番号を指定して、その選手のプレー映像（カテゴリ別のボタン＋開くと出てくる
// タイムスタンプ一覧）のHTMLを作る。カテゴリのボタンを押すと、その一番最初の
// プレーがすぐに再生される（他のプレーを見たいときは、下に出てくる一覧から選べる）。
// 試合に動画URL・時刻合わせが設定されていない場合や、選手番号が分からない場合
// （複数試合の合算表示など）は空文字を返す。
function videoClipsHtml(playerNumber) {
  const video = DATA && DATA.video;
  if (!video || !video.url || video.offsetSeconds === null || video.offsetSeconds === undefined) return '';
  if (playerNumber === null || playerNumber === undefined) return '';
  const plays = (video.plays || []).filter(p => p.number === playerNumber);
  if (!plays.length) return '';

  let html = '<div class="video-clip-groups">';
  let any = false;
  VIDEO_CLIP_CATEGORIES.forEach(cat => {
    const matched = plays.filter(p => p.skill === cat.skill && p.evaluation === cat.evaluation);
    if (!matched.length) return;
    any = true;
    const groupId = 'videoGroup_' + cat.key + '_' + playerNumber;
    const withSeconds = matched.map(p => ({
      ...p, videoSeconds: Math.max(0, Math.round(p.wallclock + video.offsetSeconds)),
    }));
    const firstSeconds = withSeconds[0].videoSeconds;
    html += `<div class="video-clip-group">`
      + `<button type="button" class="player-chip" `
      + `onclick="playVideoClip(${firstSeconds}); toggleVideoClipGroup('${groupId}')">`
      + `${cat.label}（${matched.length}）</button>`
      + `<div id="${groupId}" class="video-clip-list" style="display:none;margin-top:8px">`
      + withSeconds.map(p => {
          const mm = Math.floor(p.videoSeconds / 60);
          const ss = String(p.videoSeconds % 60).padStart(2, '0');
          return `<button type="button" class="video-clip-link" onclick="playVideoClip(${p.videoSeconds}, this)">`
            + `第${p.setNumber}セット ${mm}:${ss}</button>`;
        }).join('')
      + '</div></div>';
  });
  html += '</div>';
  return any ? html : '';
}

function toggleVideoClipGroup(id) {
  const el = document.getElementById(id);
  if (!el) return;
  el.style.display = el.style.display === 'none' ? '' : 'none';
}

// ==================== 得点パターン（どうやって得点し、どうやって失点しているか） ====================
// 得点推移(scoreProgression)のpointsには、得点/失点のたびに「決定(#)かミス(=)か」と
// 「どのスキル(サーブ/アタック/ブロック)で決まったか」が記録されている。それを集計して、
// 自チームの得点は アタック決定／サーブエース／ブロック得点／相手ミス の4種類に、
// 失点は 被決定／被サーブエース／被ブロック／自チームミス の4種類に振り分ける。
function computeScoringPattern(bySetList) {
  const counts = {
    ownKill: 0, ownAce: 0, ownBlock: 0, opponentError: 0,
    oppKill: 0, oppAce: 0, oppBlock: 0, ownError: 0,
  };
  let trueOwnTotal = 0, trueOppTotal = 0;
  (bySetList || []).forEach(s => {
    const pts = (s && s.points) || [];
    pts.forEach(pt => {
      if (pt.scoringTeam === 'own') {
        if (pt.byError) counts.opponentError += 1;
        else if (pt.skill === 'S') counts.ownAce += 1;
        else if (pt.skill === 'A') counts.ownKill += 1;
        else if (pt.skill === 'B') counts.ownBlock += 1;
      } else if (pt.scoringTeam === 'opponent') {
        if (pt.byError) counts.ownError += 1;
        else if (pt.skill === 'S') counts.oppAce += 1;
        else if (pt.skill === 'A') counts.oppKill += 1;
        else if (pt.skill === 'B') counts.oppBlock += 1;
      }
    });
    if (pts.length) {
      const last = pts[pts.length - 1];
      trueOwnTotal += last.own;
      trueOppTotal += last.opponent;
    }
  });
  return { counts, trueOwnTotal, trueOppTotal };
}

function scoringPatternHtml(pattern) {
  const c = pattern.counts;
  const ownTotal = pattern.trueOwnTotal, oppTotal = pattern.trueOppTotal;
  if (!ownTotal && !oppTotal) {
    return '<p class="hint">このデータはまだありません（この試合を追加した時点では未対応でした）。</p>';
  }
  const tile = (label, val, total) =>
    `<div class="stat-tile"><div class="k">${label}</div><div class="v">${val}<br><span style="font-size:12px;font-weight:400">${pct(dvRate(val, total))}</span></div></div>`;
  let html = `<h3 class="section-label" style="margin-top:0">得点（合計${ownTotal}点）</h3>`;
  html += '<div class="stat-tiles">'
    + tile('スパイク決定', c.ownKill, ownTotal)
    + tile('サーブエース', c.ownAce, ownTotal)
    + tile('ブロック得点', c.ownBlock, ownTotal)
    + tile('相手ミス', c.opponentError, ownTotal)
    + '</div>';
  html += `<h3 class="section-label">失点（合計${oppTotal}点）</h3>`;
  html += '<div class="stat-tiles">'
    + tile('被決定', c.oppKill, oppTotal)
    + tile('被サーブエース', c.oppAce, oppTotal)
    + tile('被ブロック', c.oppBlock, oppTotal)
    + tile('自チームミス', c.ownError, oppTotal)
    + '</div>';
  return html;
}

// ローテーションごとのサイドアウト率・ブレイク率を見て、一番目立つ差だけを
// 控えめな一文にする（本数が少なすぎるローテは、たまたまの数字になりやすいので対象外にする）
function rotationInsightText(rows) {
  const MIN = 3;
  const parts = [];
  const soRows = rows.filter(r => r.receive_total >= MIN && r.side_out_rate !== null && r.side_out_rate !== undefined);
  if (soRows.length >= 2) {
    const best = soRows.reduce((a, b) => (b.side_out_rate > a.side_out_rate ? b : a));
    const worst = soRows.reduce((a, b) => (b.side_out_rate < a.side_out_rate ? b : a));
    if (best.rotation !== worst.rotation) {
      parts.push(`サイドアウト率はS${best.rotation}が最も高く(${pct(best.side_out_rate)})、S${worst.rotation}が最も低い(${pct(worst.side_out_rate)})`);
    }
  }
  const brRows = rows.filter(r => r.serve_total >= MIN && r.break_rate !== null && r.break_rate !== undefined);
  if (brRows.length >= 2) {
    const best = brRows.reduce((a, b) => (b.break_rate > a.break_rate ? b : a));
    const worst = brRows.reduce((a, b) => (b.break_rate < a.break_rate ? b : a));
    if (best.rotation !== worst.rotation) {
      parts.push(`ブレイク率はS${best.rotation}が最も高く(${pct(best.break_rate)})、S${worst.rotation}が最も低い(${pct(worst.break_rate)})`);
    }
  }
  return parts.length ? '注目ポイント： ' + parts.join('／') : '';
}

// ==================== ローテーション別のサイドアウト率・ブレイク率 ====================
function sideOutBreakByRotationHtml(rows) {
  if (!rows.length) return '<p class="hint">このデータはまだありません（この試合を追加した時点では未対応でした）。</p>';
  let html = '<div style="overflow-x:auto"><table class="rotation-table"><thead><tr>'
    + '<th>ローテ</th><th>サーブ本数</th><th>ブレイク率</th><th>レシーブ本数</th><th>サイドアウト率</th>'
    + '</tr></thead><tbody>';
  rows.forEach(r => {
    html += `<tr><td>S${r.rotation}</td>`
      + `<td>${r.serve_total}</td>`
      + `<td>${pct(r.break_rate)}</td>`
      + `<td>${r.receive_total}</td>`
      + `<td>${pct(r.side_out_rate)}</td></tr>`;
  });
  html += '</tbody></table></div>';
  return html;
}

function showPage(n) {
  document.getElementById('page1').classList.toggle('active', n === 1);
  document.getElementById('page2').classList.toggle('active', n === 2);
  document.getElementById('tabBtn1').classList.toggle('active', n === 1);
  document.getElementById('tabBtn2').classList.toggle('active', n === 2);
}

// ==================== 折りたたみ（アコーディオン）まわりの共通処理 ====================
function toggleCollapsible(headerEl) {
  const wrap = headerEl.closest('.collapsible');
  if (wrap) wrap.classList.toggle('closed');
}

function toggleAllCollapsibles(containerEl, btnEl) {
  // ボタンの現在の表示文字で開閉方向を決める（一部だけ開いている状態でも、
  // 「すべて展開」ボタンを押したら必ず全部開く、という直感通りの動きにするため）
  const willExpand = !btnEl || btnEl.textContent.trim() !== 'すべて折りたたむ';
  containerEl.querySelectorAll('.collapsible').forEach(el => el.classList.toggle('closed', !willExpand));
  if (btnEl) btnEl.textContent = willExpand ? 'すべて折りたたむ' : 'すべて展開';
}

let selectedMatchKeys = new Set();
const matchDataCache = {};

function matchKey(entry) { return entry.dataFile; }

function fetchMatchData(entry) {
  const key = matchKey(entry);
  if (matchDataCache[key]) return Promise.resolve(matchDataCache[key]);
  return fetch(entry.dataFile).then(res => res.json()).then(data => {
    matchDataCache[key] = data;
    return data;
  });
}

function updateMatchListToggleLabel() {
  const n = selectedMatchKeys.size;
  document.getElementById('matchListToggleLabel').textContent =
    n > 0 ? `試合（${n}試合選択中）` : '試合';
  document.getElementById('matchComparisonHint').style.display = n >= 2 ? 'none' : '';
}

// matchDateは保存形式(DD/MM/YYYY)なので、年月でグループ化するためのキーを作る
function matchMonthKey(matchDate) {
  const parts = matchDate.split('/');
  if (parts.length === 3) {
    const [, month, year] = parts;
    return `${year}/${month.padStart(2, '0')}`;
  }
  return matchDate;
}

function renderMatchChecklist() {
  const list = document.getElementById('matchChecklist');
  list.innerHTML = '';
  const groupOrder = [];
  const groups = new Map();
  MATCHES.forEach(m => {
    const key = matchMonthKey(m.matchDate);
    if (!groups.has(key)) { groups.set(key, []); groupOrder.push(key); }
    groups.get(key).push(m);
  });

  groupOrder.forEach(key => {
    const groupMatches = groups.get(key);
    const [year, month] = key.split('/');

    const header = document.createElement('div');
    header.className = 'match-month-header';
    const label = document.createElement('span');
    label.textContent = `${year}年${parseInt(month, 10)}月（${groupMatches.length}試合）`;
    const selectBtn = document.createElement('button');
    selectBtn.type = 'button';
    selectBtn.className = 'match-month-select-btn';
    const allSelected = groupMatches.every(m => selectedMatchKeys.has(matchKey(m)));
    selectBtn.textContent = allSelected ? 'まとめて解除' : 'まとめて選択';
    selectBtn.addEventListener('click', () => {
      const nowAllSelected = groupMatches.every(m => selectedMatchKeys.has(matchKey(m)));
      groupMatches.forEach(m => {
        if (nowAllSelected) selectedMatchKeys.delete(matchKey(m));
        else selectedMatchKeys.add(matchKey(m));
      });
      updateMatchListToggleLabel();
      clearDropZoneStatus();
      renderMatchChecklist();
      renderMatchSelection();
    });
    header.appendChild(label);
    header.appendChild(selectBtn);
    list.appendChild(header);

    groupMatches.forEach(m => {
      const row = document.createElement('div');
      row.className = 'match-row';

      const itemLabel = document.createElement('label');
      itemLabel.className = 'player-check-item';
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = selectedMatchKeys.has(matchKey(m));
      cb.addEventListener('change', () => toggleMatchCheckbox(m, cb));
      const span = document.createElement('span');
      span.textContent = `${formatMatchDate(m.matchDate)} 対 ${m.opponent}`;
      itemLabel.appendChild(cb);
      itemLabel.appendChild(span);

      // チェックを入れてダッシュボード全体を読み込まなくても、
      // このアイコンを押すだけですぐその試合のサマリー画像が見られる。
      const summaryIconBtn = document.createElement('button');
      summaryIconBtn.type = 'button';
      summaryIconBtn.className = 'match-summary-icon-btn';
      summaryIconBtn.title = 'この試合のサマリーを1枚で見る';
      summaryIconBtn.setAttribute('aria-label', 'この試合のサマリーを1枚で見る');
      summaryIconBtn.textContent = '🖼';
      summaryIconBtn.addEventListener('click', () => openMatchSummaryForEntry(m));

      row.appendChild(itemLabel);
      row.appendChild(summaryIconBtn);
      list.appendChild(row);
    });
  });
  updateMatchListToggleLabel();
}

// 対戦相手を選ぶだけで、その相手との試合を全部まとめて選択する（スカウティング用のショートカット）
function renderOpponentQuickSelect() {
  const wrap = document.getElementById('opponentQuickSelectWrap');
  const select = document.getElementById('opponentQuickSelect');
  const counts = new Map();
  MATCHES.forEach(m => counts.set(m.opponent, (counts.get(m.opponent) || 0) + 1));
  const names = Array.from(counts.keys()).sort((a, b) => counts.get(b) - counts.get(a));
  if (!names.length) { wrap.style.display = 'none'; return; }
  wrap.style.display = 'block';
  select.innerHTML = '<option value="">対戦相手を選択…</option>'
    + names.map(name => `<option value="${escapeHtml(name)}">${escapeHtml(name)}（${counts.get(name)}試合）</option>`).join('');
}

function setupOpponentQuickSelect() {
  const select = document.getElementById('opponentQuickSelect');
  select.addEventListener('change', () => {
    const name = select.value;
    if (!name) return;
    selectedMatchKeys.clear();
    MATCHES.filter(m => m.opponent === name).forEach(m => selectedMatchKeys.add(matchKey(m)));
    updateMatchListToggleLabel();
    clearDropZoneStatus();
    renderMatchChecklist();
    renderMatchSelection();
  });
}

function setupMatchListToggle() {
  const btn = document.getElementById('matchListToggleBtn');
  const list = document.getElementById('matchChecklist');
  const arrow = document.getElementById('matchListToggleArrow');
  btn.addEventListener('click', () => {
    const isOpen = list.style.display !== 'none';
    list.style.display = isOpen ? 'none' : 'flex';
    arrow.textContent = isOpen ? '▾' : '▴';
  });
}

function toggleMatchCheckbox(entry, cbEl) {
  const key = matchKey(entry);
  if (cbEl.checked) {
    selectedMatchKeys.add(key);
  } else {
    selectedMatchKeys.delete(key);
  }
  updateMatchListToggleLabel();
  clearDropZoneStatus();
  renderMatchChecklist();
  renderMatchSelection();
}

function clearDropZoneStatus() {
  const statusEl = document.getElementById('dropZoneStatus');
  if (statusEl) statusEl.innerHTML = '';
}

function renderMatchSelection() {
  const entries = Array.from(selectedMatchKeys)
    .map(k => MATCHES.find(m => matchKey(m) === k))
    .filter(Boolean);

  if (entries.length === 0) {
    document.getElementById('mainArea').style.display = 'none';
    document.getElementById('playerSelectCard').style.display = 'none';
    document.getElementById('matchComparisonCard').style.display = 'none';
    return;
  }

  document.getElementById('playerSelectCard').style.display = 'block';
  renderTeamGrid();
  renderPlayerGrid();

  if (entries.length >= 2) {
    Promise.all(entries.map(fetchMatchData)).then(dataList => {
      // 非同期で待っている間に試合の選択が変わっていたら、古い結果は描画しない
      const stillSame = Array.from(selectedMatchKeys).sort().join(',') === entries.map(matchKey).sort().join(',');
      if (!stillSame) return;
      showSingleMatch(dataList[0], entries, dataList);
      showMatchComparison(dataList);
    });
  } else {
    fetchMatchData(entries[0]).then(data => {
      const stillSame = selectedMatchKeys.size === 1 && selectedMatchKeys.has(matchKey(entries[0]));
      if (!stillSame) return;
      showSingleMatch(data);
    });
    document.getElementById('matchComparisonCard').style.display = 'none';
  }
}

function showSingleMatch(data, entries, dataList) {
  DATA = data;
  document.getElementById('mainArea').style.display = 'block';
  document.getElementById('matchTitle').textContent =
    `${DATA.teamLabel} 対 ${DATA.opponent}（${formatMatchDate(DATA.matchDate)}）`;
  const combined = entries && dataList && entries.length >= 2;
  if (combined) {
    showRotationTablesCombined(entries, dataList);
    showOpponentCourseSetsCombined(entries, dataList);
    showSideOutBreakTableCombined(entries, dataList);
    showScoringPatternCombined(entries, dataList);
    showScoreProgressionChartsCombined(entries, dataList);
    showRunHighlightsCombined(entries, dataList);
  } else {
    showRotationTables();
    showOpponentCourseSets();
    showSideOutBreakTable();
    showScoringPattern();
    showScoreProgressionCharts();
    showRunHighlights();
  }
  showPage(1);
  renderSingleMatchPlayerView();
  if (combined) {
    document.getElementById('startingLineupSetButtons').innerHTML = '';
    document.getElementById('startingLineupBody').innerHTML =
      '<p class="hint">スタメンは1試合を選んでいるときに表示されます。</p>';
    document.getElementById('leaderboardLabel').textContent = `選手ランキング（選んだ${entries.length}試合の合計）`;
    document.getElementById('leaderboardBody').innerHTML = playerLeaderboardHtml(combinedPlayerTotalsList(dataList));
  } else {
    showStartingLineups(DATA.scoreProgression.bySet);
    document.getElementById('leaderboardLabel').textContent = '選手ランキング（この試合）';
    document.getElementById('leaderboardBody').innerHTML = playerLeaderboardHtml(DATA.total || []);
  }
  // サマリー画像ボタンは「1試合だけを選んでいる」ときだけ表示する
  // （複数試合の合算では「その試合の」サマリーという意味が成り立たないため）。
  document.getElementById('summaryBtn').style.display = combined ? 'none' : 'inline-flex';
}

// ==================== 試合サマリー（1枚まとめ画像） ====================
// 選手・チームのスタッツをLINEなどでそのまま共有できる、1枚の画像を <canvas> に直接描画する。
// 外部ライブラリ（html2canvas等）には頼らず自前で描く方式にしている
// （体育館などネット環境が不安定な場所でも、外部CDNの読み込み待ちや失敗なしに確実に保存できるように）。
// 色はダーク/ライトモードの影響を受けないよう、CSS変数を使わずすべて直接指定している。
function summaryRoundRect(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}

function summaryEllipsize(ctx, text, maxWidth) {
  if (ctx.measureText(text).width <= maxWidth) return text;
  let s = text;
  while (s.length > 1 && ctx.measureText(s + '…').width > maxWidth) {
    s = s.slice(0, -1);
  }
  return s + '…';
}

function summarySetScoreLabels(scoreProgression) {
  return (scoreProgression.bySet || []).map(s => {
    const last = s.points[s.points.length - 1];
    if (!last) return null;
    return { label: `${last.own}-${last.opponent}`, win: last.own > last.opponent };
  }).filter(Boolean);
}

// setIndex が null なら「全セット合計」、数値ならその番号のセット（0始まり）だけの
// チーム/選手スタッツを取り出す。決定率・AB率などの各種レートはdvPlayerSummary側で
// 常に計算済みなので、対象を選手全員(total)からそのセットの選手(sets[i].players)に
// 差し替えるだけで、セットごとの決定率・効果率・AB率などがそのまま出せる。
function summaryScope(data, setIndex) {
  if (setIndex === null || setIndex === undefined) {
    return {
      label: '全セット合計',
      team: data.team.total,
      sideOut: data.sideOutBreak.total,
      players: data.total || [],
    };
  }
  return {
    label: `第${setIndex + 1}セット`,
    team: (data.team.sets || [])[setIndex] || {},
    sideOut: (data.sideOutBreak.bySet || [])[setIndex] || {},
    players: ((data.sets || [])[setIndex] || {}).players || [],
  };
}

function summaryTeamTiles(scope) {
  const t = scope.team;
  const sob = scope.sideOut;
  return [
    { k: 'スパイク決定率', v: pct(t.kill_rate), sub: `${t.attempts || 0}本中${t.points || 0}本決定` },
    { k: 'サーブ効果率', v: pct(t.serve_efficiency), sub: `${t.serve_attempts || 0}本中${(t.serve_aces || 0) + (t.serve_half_credit || 0)}本` },
    { k: 'レシーブAB率', v: pct(t.receive_ab_rate), sub: `${t.receive_attempts || 0}本中${(t.receive_a || 0) + (t.receive_b || 0)}本` },
    { k: 'ブロック得点', v: `${t.block_points || 0}本`, sub: `被ブロック${t.blocked || 0}本` },
    { k: 'サイドアウト率', v: pct(sob.side_out_rate), sub: `レシーブ${sob.receive_total || 0}本` },
    { k: 'ブレイク率', v: pct(sob.break_rate), sub: `サーブ${sob.serve_total || 0}本` },
  ];
}

function summaryPlayerRows(scope) {
  return (scope.players || [])
    .filter(p => (p.attempts || 0) + (p.serve_attempts || 0) + (p.receive_attempts || 0) + (p.block_attempts || 0) > 0)
    .sort((a, b) => a.number - b.number);
}

// data(試合1件分のJSON)から、共有用の1枚画像を<canvas>に描いて返す。
// setIndex が null なら全セット合計、数値ならそのセット（0始まり）だけを表示する。
function drawMatchSummaryCanvas(canvas, data, setIndex) {
  const W = 720;
  const PAD = 26;
  const SCALE = 2; // 保存した画像がぼやけないよう、2倍の解像度で描く

  const result = computeSetScoreSummary(data.scoreProgression);
  const setScores = summarySetScoreLabels(data.scoreProgression);
  const scope = summaryScope(data, setIndex);
  const tiles = summaryTeamTiles(scope);
  const players = summaryPlayerRows(scope);

  // --- 高さを先に計算する ---
  const headerTopH = 92;   // eyebrow + 対戦相手名 + 日付
  const headerResultH = 58; // 勝敗バッジ + セットスコア行
  const headerH = headerTopH + headerResultH;
  const bodyTop = headerH + 26;
  const tileRowH = 62, tileGap = 10;
  const tileRows = Math.ceil(tiles.length / 3);
  const tilesH = tileRows * tileRowH + (tileRows - 1) * tileGap;
  const tableTop = bodyTop + 20 + tilesH + 30 + 20;
  const rowH = 36, headRowH = 26;
  const tableBodyH = Math.max(players.length, 1) * rowH;
  const footH = 40;
  const H = tableTop + headRowH + tableBodyH + footH;

  canvas.width = W * SCALE;
  canvas.height = H * SCALE;
  canvas.style.width = W + 'px';
  const ctx = canvas.getContext('2d');
  ctx.setTransform(SCALE, 0, 0, SCALE, 0, 0);
  ctx.textBaseline = 'alphabetic';
  ctx.font = '13px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';

  // 背景
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, W, H);

  // ヘッダー（グラデーション帯）
  const grad = ctx.createLinearGradient(0, 0, W, headerH);
  grad.addColorStop(0, '#16357a');
  grad.addColorStop(1, '#1d5fc4');
  summaryRoundRect(ctx, 0, 0, W, headerH, 10);
  ctx.save(); ctx.clip();
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, W, headerH);
  ctx.restore();

  ctx.fillStyle = 'rgba(255,255,255,0.85)';
  ctx.font = '600 12px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
  ctx.fillText(`${data.teamLabel}　試合サマリー`, PAD, 36);

  ctx.fillStyle = '#ffffff';
  ctx.font = '800 25px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
  ctx.fillText(summaryEllipsize(ctx, `対 ${data.opponent}`, W - PAD * 2), PAD, 66);

  ctx.fillStyle = 'rgba(255,255,255,0.85)';
  ctx.font = '13px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
  ctx.fillText(formatMatchDate(data.matchDate), PAD, 86);

  // 勝敗バッジ
  const badgeText = `${result.setsWon}勝${result.setsLost}敗`;
  ctx.font = '800 19px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
  const badgeTextW = ctx.measureText(badgeText).width;
  const badgeW = badgeTextW + 32, badgeH = 32, badgeY = headerTopH + 8;
  ctx.fillStyle = 'rgba(255,255,255,0.16)';
  summaryRoundRect(ctx, PAD, badgeY, badgeW, badgeH, badgeH / 2);
  ctx.fill();
  ctx.fillStyle = '#ffffff';
  ctx.textBaseline = 'middle';
  ctx.fillText(badgeText, PAD + 16, badgeY + badgeH / 2 + 1);
  ctx.textBaseline = 'alphabetic';

  // セットごとのスコア（勝ったセットは白背景で強調。今表示中のセットには白い枠を付ける）
  let sx = PAD + badgeW + 12;
  const pillY = badgeY, pillH = badgeH;
  ctx.font = '700 12px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
  setScores.forEach((s, i) => {
    const tw = ctx.measureText(s.label).width;
    const pw = tw + 16;
    if (sx + pw > W - PAD) return; // 万一入りきらない場合は打ち切り（通常5セット程度なら収まる）
    const py = pillY + (pillH - 24) / 2;
    ctx.fillStyle = s.win ? 'rgba(255,255,255,0.92)' : 'rgba(255,255,255,0.14)';
    summaryRoundRect(ctx, sx, py, pw, 24, 6);
    ctx.fill();
    if (i === setIndex) {
      ctx.strokeStyle = '#ffffff';
      ctx.lineWidth = 2;
      summaryRoundRect(ctx, sx + 1, py + 1, pw - 2, 22, 5);
      ctx.stroke();
    }
    ctx.fillStyle = s.win ? '#16357a' : '#ffffff';
    ctx.textBaseline = 'middle';
    ctx.fillText(s.label, sx + 8, pillY + pillH / 2 + 1);
    ctx.textBaseline = 'alphabetic';
    sx += pw + 8;
  });

  // 本体の枠線
  ctx.strokeStyle = '#e4e4e0';
  ctx.lineWidth = 1;
  summaryRoundRect(ctx, 0.5, headerH + 0.5, W - 1, H - headerH - 1, 10);
  ctx.stroke();

  // セクション見出し
  function sectionTitle(text, y) {
    ctx.fillStyle = '#8a93a3';
    ctx.font = '700 12px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
    ctx.fillText(text, PAD, y);
  }
  sectionTitle(`チームスタッツ（${scope.label}）`, bodyTop);

  // チームスタッツ タイル
  const tileW = (W - PAD * 2 - tileGap * 2) / 3;
  tiles.forEach((tl, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const tx = PAD + col * (tileW + tileGap);
    const ty = bodyTop + 14 + row * (tileRowH + tileGap);
    ctx.fillStyle = '#f6f7fa';
    ctx.strokeStyle = '#e4e4e0';
    summaryRoundRect(ctx, tx, ty, tileW, tileRowH, 8);
    ctx.fill(); ctx.stroke();

    ctx.fillStyle = '#6b7280';
    ctx.font = '11px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
    ctx.fillText(summaryEllipsize(ctx, tl.k, tileW - 16), tx + 10, ty + 20);

    ctx.fillStyle = '#16357a';
    ctx.font = '800 18px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
    ctx.fillText(tl.v, tx + 10, ty + 42);

    ctx.fillStyle = '#9aa1ae';
    ctx.font = '10px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
    ctx.fillText(summaryEllipsize(ctx, tl.sub, tileW - 16), tx + 10, ty + 55);
  });

  const playerSectionY = bodyTop + 14 + tilesH + 30;
  sectionTitle(`選手別スタッツ（${scope.label}）`, playerSectionY);

  // 選手テーブル
  const cols = [
    { key: 'num', label: '', x: PAD, w: 28, align: 'left' },
    { key: 'name', label: '選手', x: PAD + 28, w: 78, align: 'left' },
    { key: 'spike', label: 'スパイク', x: 0, w: 0, align: 'center' },
    { key: 'serve', label: 'サーブ', x: 0, w: 0, align: 'center' },
    { key: 'receive', label: 'レシーブ', x: 0, w: 0, align: 'center' },
    { key: 'block', label: 'ブロック', x: 0, w: 0, align: 'center' },
  ];
  const fixedW = 28 + 78;
  const flexW = (W - PAD * 2 - fixedW) / 4;
  let cx = PAD + fixedW;
  for (let i = 2; i < cols.length; i++) { cols[i].x = cx; cols[i].w = flexW; cx += flexW; }

  const theadY = tableTop;
  ctx.fillStyle = '#6b7280';
  ctx.font = '700 10.5px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
  cols.forEach(c => {
    if (!c.label) return;
    const textW = ctx.measureText(c.label).width;
    const tx = c.align === 'left' ? c.x : c.x + c.w / 2 - textW / 2;
    ctx.fillText(c.label, tx, theadY + headRowH - 9);
  });
  ctx.strokeStyle = '#e4e4e0';
  ctx.lineWidth = 1.5;
  ctx.beginPath();
  ctx.moveTo(PAD, theadY + headRowH);
  ctx.lineTo(W - PAD, theadY + headRowH);
  ctx.stroke();

  if (!players.length) {
    ctx.fillStyle = '#9aa1ae';
    ctx.font = '13px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
    ctx.fillText('出場記録がありません。', PAD, theadY + headRowH + 22);
  }

  players.forEach((p, i) => {
    const ry = theadY + headRowH + i * rowH;
    if (i % 2 === 1) {
      ctx.fillStyle = '#f9fafc';
      ctx.fillRect(PAD, ry, W - PAD * 2, rowH);
    }
    ctx.strokeStyle = '#eef0f3';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(PAD, ry + rowH);
    ctx.lineTo(W - PAD, ry + rowH);
    ctx.stroke();

    const midY = ry + 16, subY = ry + 29;
    ctx.fillStyle = '#9aa1ae';
    ctx.font = '700 12px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
    ctx.fillText(String(p.number), cols[0].x, midY + 4);

    ctx.fillStyle = '#1b2333';
    ctx.font = '700 13px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
    ctx.fillText(summaryEllipsize(ctx, p.name, cols[1].w), cols[1].x, midY + 4);

    function cell(col, main, sub) {
      ctx.fillStyle = '#1b2333';
      ctx.font = '700 12.5px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
      let tw = ctx.measureText(main).width;
      ctx.fillText(main, col.x + col.w / 2 - tw / 2, midY + 4);
      if (sub) {
        ctx.fillStyle = '#9aa1ae';
        ctx.font = '10px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
        tw = ctx.measureText(sub).width;
        ctx.fillText(sub, col.x + col.w / 2 - tw / 2, subY + 4);
      }
    }
    // ％【本数】の1行表示（サイトの表と同じ形式）。mainは太字、bracketは小さいグレー文字。
    function cellRate(col, main, bracket) {
      ctx.font = '700 12.5px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
      const mainW = ctx.measureText(main).width;
      let bracketW = 0;
      if (bracket) {
        ctx.font = '10px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
        bracketW = ctx.measureText(bracket).width;
      }
      const gap = bracket ? 3 : 0;
      let x = col.x + col.w / 2 - (mainW + gap + bracketW) / 2;
      ctx.fillStyle = '#1b2333';
      ctx.font = '700 12.5px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
      ctx.fillText(main, x, midY + 4);
      if (bracket) {
        x += mainW + gap;
        ctx.fillStyle = '#9aa1ae';
        ctx.font = '10px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
        ctx.fillText(bracket, x, midY + 4);
      }
    }
    const killBracket = p.attempts ? `【${p.points}/${p.attempts}本】` : '';
    cellRate(cols[2], `決定率${pct(p.kill_rate)}`, killBracket);
    const serveBracket = p.serve_attempts ? `【${(p.serve_aces || 0) + (p.serve_half_credit || 0)}/${p.serve_attempts}本】` : '';
    cellRate(cols[3], `効果率${pct(p.serve_efficiency)}`, serveBracket);
    const receiveBracket = p.receive_attempts ? `【${(p.receive_a || 0) + (p.receive_b || 0)}/${p.receive_attempts}本】` : '';
    cellRate(cols[4], `AB率${pct(p.receive_ab_rate)}`, receiveBracket);
    cell(cols[5], `${p.block_points}本`, '');
  });

  // フッター
  ctx.fillStyle = '#9aa1ae';
  ctx.font = '10.5px "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif';
  const footText = `${data.teamLabel}　スタッツ分析ツール`;
  const footW = ctx.measureText(footText).width;
  ctx.fillText(footText, W / 2 - footW / 2, H - 16);
}

let summaryImageDataUrl = null;
let summaryData = null; // 今モーダルに表示している試合のデータ（メイン画面のDATAとは独立）
let summarySetIndex = null; // null=全セット合計、0以上ならそのセット（0始まり）
const summaryCanvasEl = document.createElement('canvas'); // 画面には出さず、描画専用に使う

// 「全セット」「第1セット」…のタブを、今の試合のセット数に合わせて作り直す
function renderSummarySetTabs(data) {
  const wrap = document.getElementById('summarySetTabs');
  wrap.innerHTML = '';
  const setCount = (data.scoreProgression.bySet || []).length;
  const options = [{ index: null, label: '全セット' }];
  for (let i = 0; i < setCount; i++) options.push({ index: i, label: `第${i + 1}セット` });
  options.forEach(opt => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'summary-set-tab' + (summarySetIndex === opt.index ? ' active' : '');
    btn.textContent = opt.label;
    btn.addEventListener('click', () => {
      if (summarySetIndex === opt.index) return;
      summarySetIndex = opt.index;
      renderSummarySetTabs(data);
      renderSummaryImage();
    });
    wrap.appendChild(btn);
  });
}

// 今の summaryData / summarySetIndex の組み合わせで画像を描き直す
function renderSummaryImage() {
  const data = summaryData;
  if (!data) return;
  const img = document.getElementById('summaryImg');
  const loading = document.getElementById('summaryLoading');
  img.style.display = 'none';
  loading.style.display = 'block';
  const setIndexAtDrawTime = summarySetIndex;
  const draw = () => {
    // 描画中に別の試合・別のセットが選ばれていたら、古い結果は捨てて何もしない
    if (summaryData !== data || summarySetIndex !== setIndexAtDrawTime) return;
    drawMatchSummaryCanvas(summaryCanvasEl, data, setIndexAtDrawTime);
    summaryImageDataUrl = summaryCanvasEl.toDataURL('image/png');
    img.src = summaryImageDataUrl;
    loading.style.display = 'none';
    img.style.display = 'inline-block';
  };
  // フォントの読み込みを待ってから描くと、初回表示でも文字化けやガタつきが起きにくい
  if (document.fonts && document.fonts.ready) {
    document.fonts.ready.then(draw).catch(draw);
  } else {
    draw();
  }
}

// data(その試合のJSON)をもとにモーダルを開いて描画する。
// メイン画面でその試合を選んでいなくても、試合一覧のアイコンから直接呼べるようにしている。
function openMatchSummaryWithData(data) {
  summaryData = data;
  summarySetIndex = null; // 開くたびに「全セット」から始める
  document.getElementById('summaryOverlay').style.display = 'flex';
  renderSummarySetTabs(data);
  renderSummaryImage();
}

// メイン画面で今表示している試合（DATA）のサマリーを開く
function openMatchSummary() {
  if (!DATA) return;
  openMatchSummaryWithData(DATA);
}

// 試合一覧のアイコンから、その場でその試合のサマリーを開く
// （チェックを入れてダッシュボード全体を読み込まなくても済むようにするため）
function openMatchSummaryForEntry(entry) {
  fetchMatchData(entry).then(data => openMatchSummaryWithData(data));
}

function closeMatchSummary() {
  document.getElementById('summaryOverlay').style.display = 'none';
}

function saveMatchSummaryImage() {
  if (!summaryImageDataUrl || !summaryData) return;
  const dateLabel = (summaryData.matchDate || '').replace(/[^0-9]/g, '') || 'match';
  const setLabel = summarySetIndex === null ? '' : `_第${summarySetIndex + 1}セット`;
  const link = document.createElement('a');
  link.download = `${dateLabel}_${summaryData.opponent}${setLabel}_サマリー.png`;
  link.href = summaryImageDataUrl;
  document.body.appendChild(link);
  link.click();
  link.remove();
}

document.getElementById('summaryBtn').addEventListener('click', openMatchSummary);
document.getElementById('summaryCloseBtn').addEventListener('click', closeMatchSummary);
document.getElementById('summarySaveBtn').addEventListener('click', saveMatchSummaryImage);
document.getElementById('summaryOverlay').addEventListener('click', (e) => {
  if (e.target.id === 'summaryOverlay') closeMatchSummary();
});

// ==================== 選手ランキング（規定本数以上の選手が対象） ====================
function playerLeaderboardHtml(players) {
  const MIN_ATTEMPTS = 3;
  function topList(list, key, minKey) {
    return list
      .filter(p => (p[minKey] || 0) >= MIN_ATTEMPTS && p[key] !== null && p[key] !== undefined)
      .sort((a, b) => b[key] - a[key])
      .slice(0, 5);
  }
  function rowsHtml(list, key, numGetter, denKey) {
    if (!list.length) return '<p class="hint">規定本数（' + MIN_ATTEMPTS + '本）以上の選手がいません。</p>';
    return '<table class="rotation-table"><thead><tr><th>順位</th><th>選手</th><th>値</th></tr></thead><tbody>'
      + list.map((p, i) => `<tr><td>${i + 1}</td><td>${escapeHtml(p.name)}</td><td>${pctCountBlock(p[key], numGetter(p), p[denKey])}</td></tr>`).join('')
      + '</tbody></table>';
  }
  const byKill = topList(players, 'kill_rate', 'attempts');
  const byServe = topList(players, 'serve_efficiency', 'serve_attempts');
  const byReceive = topList(players, 'receive_a_rate', 'receive_attempts');
  const byReceiveAB = topList(players, 'receive_ab_rate', 'receive_attempts');

  return '<div class="rank-grid">'
    + `<div><h4>スパイク決定率</h4>${rowsHtml(byKill, 'kill_rate', p => p.points, 'attempts')}</div>`
    + `<div><h4>サーブ効果率</h4>${rowsHtml(byServe, 'serve_efficiency', p => p.serve_aces + p.serve_half_credit, 'serve_attempts')}</div>`
    + `<div><h4>レシーブA率</h4>${rowsHtml(byReceive, 'receive_a_rate', p => p.receive_a, 'receive_attempts')}</div>`
    + `<div><h4>レシーブAB率</h4>${rowsHtml(byReceiveAB, 'receive_ab_rate', p => p.receive_a + p.receive_b, 'receive_attempts')}</div>`
    + '</div>';
}

function computeSetScoreSummary(scoreProgression) {
  let setsWon = 0, setsLost = 0;
  (scoreProgression.bySet || []).forEach(s => {
    const last = s.points[s.points.length - 1];
    if (!last) return;
    if (last.own > last.opponent) setsWon++;
    else if (last.opponent > last.own) setsLost++;
  });
  return { setsWon, setsLost };
}

function matchComparisonRowHtml(label, matches, getter, isPct) {
  let html = `<tr><td>${escapeHtml(label)}</td>`;
  matches.forEach(m => {
    const v = getter(m);
    html += `<td>${isPct ? pct(v) : v}</td>`;
  });
  html += '</tr>';
  return html;
}

function matchComparisonSectionHtml(title, colspan) {
  return `<tr class="section-row"><td colspan="${colspan}">${escapeHtml(title)}</td></tr>`;
}

function matchComparisonTableHtml(matches) {
  const colspan = matches.length + 1;
  let html = '<div style="overflow-x:auto"><table class="rotation-table comparison-table">';
  html += '<thead><tr><th>項目</th>';
  matches.forEach(m => {
    html += `<th>${escapeHtml(m.opponent)}<br><span style="font-weight:400;font-size:12px">${escapeHtml(formatMatchDate(m.matchDate))}</span></th>`;
  });
  html += '</tr></thead><tbody>';

  html += matchComparisonSectionHtml('試合結果', colspan);
  html += '<tr><td>セット</td>';
  matches.forEach(m => {
    const s = computeSetScoreSummary(m.scoreProgression);
    html += `<td>${s.setsWon}勝${s.setsLost}敗</td>`;
  });
  html += '</tr>';

  html += matchComparisonSectionHtml('スパイク（全セット合計）', colspan);
  html += matchComparisonRowHtml('打数', matches, m => m.team.total.attempts, false);
  html += matchComparisonRowHtml('得点', matches, m => m.team.total.points, false);
  html += matchComparisonRowHtml('ミス', matches, m => m.team.total.errors, false);
  html += matchComparisonRowHtml('決定率', matches, m => m.team.total.kill_rate, true);
  html += matchComparisonRowHtml('効果率', matches, m => m.team.total.efficiency, true);

  html += matchComparisonSectionHtml('ブロック（全セット合計）', colspan);
  html += matchComparisonRowHtml('本数', matches, m => m.team.total.block_attempts, false);
  html += matchComparisonRowHtml('得点', matches, m => m.team.total.block_points, false);
  html += matchComparisonRowHtml('ワンチ', matches, m => m.team.total.block_touch_own + m.team.total.block_touch_opp, false);
  html += matchComparisonRowHtml('失点', matches, m => m.team.total.block_errors, false);

  html += matchComparisonSectionHtml('サーブ（全セット合計）', colspan);
  html += matchComparisonRowHtml('打数', matches, m => m.team.total.serve_attempts, false);
  html += matchComparisonRowHtml('エース', matches, m => m.team.total.serve_aces, false);
  html += matchComparisonRowHtml('ミス', matches, m => m.team.total.serve_errors, false);
  html += matchComparisonRowHtml('効果本数', matches, m => m.team.total.serve_half_credit, false);
  html += matchComparisonRowHtml('エース率', matches, m => m.team.total.serve_ace_rate, true);
  html += matchComparisonRowHtml('ミス率', matches, m => m.team.total.serve_error_rate, true);
  html += matchComparisonRowHtml('効果率', matches, m => m.team.total.serve_efficiency, true);

  html += matchComparisonSectionHtml('レシーブ（全セット合計）', colspan);
  html += matchComparisonRowHtml('本数', matches, m => m.team.total.receive_attempts, false);
  html += matchComparisonRowHtml('Aパス', matches, m => m.team.total.receive_a, false);
  html += matchComparisonRowHtml('Bパス', matches, m => m.team.total.receive_b, false);
  html += matchComparisonRowHtml('Cパス', matches, m => m.team.total.receive_c, false);
  html += matchComparisonRowHtml('Dパス', matches, m => m.team.total.receive_d, false);
  html += matchComparisonRowHtml('ミス', matches, m => m.team.total.receive_errors, false);
  html += matchComparisonRowHtml('返球率', matches, m => m.team.total.receive_return_rate, true);
  html += matchComparisonRowHtml('A率', matches, m => m.team.total.receive_a_rate, true);
  html += matchComparisonRowHtml('AB率', matches, m => m.team.total.receive_ab_rate, true);

  html += matchComparisonSectionHtml('サイドアウト率・ブレイク率（チーム全体）', colspan);
  html += matchComparisonRowHtml('サーブ本数', matches, m => m.sideOutBreak.total.serve_total, false);
  html += matchComparisonRowHtml('ブレイク率', matches, m => m.sideOutBreak.total.break_rate, true);
  html += matchComparisonRowHtml('レシーブ本数', matches, m => m.sideOutBreak.total.receive_total, false);
  html += matchComparisonRowHtml('サイドアウト率', matches, m => m.sideOutBreak.total.side_out_rate, true);

  html += '</tbody></table></div>';
  return html;
}

// ==================== 選手個人の、試合をまたいだ推移グラフ（決定率・効果率） ====================
// 調子の波が一目でわかるように、選んだ試合を日付順に並べて折れ線グラフにする。
// 出場していない試合（記録なし）は線をつながず、点だけ飛ばす。
function playerTrendChartSvg(number, matches) {
  const sorted = matches.slice().sort((a, b) => dvDateSlug(a.matchDate).localeCompare(dvDateSlug(b.matchDate)));
  const points = sorted.map(m => {
    const p = (m.total || []).find(pp => pp.number === number);
    const displayDate = formatMatchDate(m.matchDate);
    const parts = displayDate.split('/');
    const shortLabel = parts.length === 3 ? `${parts[1]}/${parts[2]}` : displayDate;
    return {
      fullLabel: `${displayDate} 対 ${m.opponent}`,
      shortLabel,
      kill_rate: p ? p.kill_rate : null,
      efficiency: p ? p.efficiency : null,
    };
  });
  const values = [];
  points.forEach(pt => {
    if (pt.kill_rate !== null && pt.kill_rate !== undefined) values.push(pt.kill_rate);
    if (pt.efficiency !== null && pt.efficiency !== undefined) values.push(pt.efficiency);
  });
  if (values.length < 2) return '<p class="hint">推移グラフを表示するには、出場記録のある試合が2試合以上必要です。</p>';

  const W = 640, H = 220;
  const padL = 42, padR = 16, padT = 14, padB = 30;
  const plotW = W - padL - padR, plotH = H - padT - padB;
  const n = points.length;

  const rawMin = Math.min(0, ...values);
  const rawMax = Math.max(0, ...values);
  const span = Math.max(0.1, rawMax - rawMin);
  const yMin = rawMin - span * 0.15;
  const yMax = rawMax + span * 0.15;

  const xAt = i => padL + (n <= 1 ? plotW / 2 : (i / (n - 1)) * plotW);
  const yAt = v => padT + plotH - ((v - yMin) / (yMax - yMin)) * plotH;

  let gridSvg = `<line x1="${padL}" y1="${padT}" x2="${padL}" y2="${(H - padB).toFixed(1)}" stroke="var(--grid)" stroke-width="1"/>`;
  if (yMin < 0 && yMax > 0) {
    const zeroY = yAt(0).toFixed(1);
    gridSvg += `<line x1="${padL}" y1="${zeroY}" x2="${W - padR}" y2="${zeroY}" stroke="var(--grid)" stroke-width="1"/>`;
  }
  gridSvg += `<text x="${(padL - 6).toFixed(1)}" y="${(yAt(yMax) + 4).toFixed(1)}" text-anchor="end" font-size="11" fill="var(--text-muted)">${pct(yMax)}</text>`;
  gridSvg += `<text x="${(padL - 6).toFixed(1)}" y="${(yAt(yMin) + 4).toFixed(1)}" text-anchor="end" font-size="11" fill="var(--text-muted)">${pct(yMin)}</text>`;

  function seriesSvg(key, colorVar) {
    let segHtml = '';
    let seg = [];
    points.forEach((pt, i) => {
      const v = pt[key];
      if (v === null || v === undefined) {
        if (seg.length > 1) {
          segHtml += `<polyline points="${seg.join(' ')}" fill="none" stroke="${colorVar}" `
            + `stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
        }
        seg = [];
        return;
      }
      seg.push(`${xAt(i).toFixed(1)},${yAt(v).toFixed(1)}`);
    });
    if (seg.length > 1) {
      segHtml += `<polyline points="${seg.join(' ')}" fill="none" stroke="${colorVar}" `
        + `stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
    }
    const dots = points.map((pt, i) => {
      const v = pt[key];
      if (v === null || v === undefined) return '';
      return `<circle cx="${xAt(i).toFixed(1)}" cy="${yAt(v).toFixed(1)}" r="3.5" fill="${colorVar}" `
        + `stroke="var(--surface-1)" stroke-width="1"><title>${escapeHtml(pt.fullLabel)}: ${pct(v)}</title></circle>`;
    }).join('');
    return segHtml + dots;
  }

  const xLabels = points.map((pt, i) => {
    const anchor = i === 0 ? 'start' : (i === n - 1 ? 'end' : 'middle');
    return `<text x="${xAt(i).toFixed(1)}" y="${H - 8}" text-anchor="${anchor}" font-size="10.5" `
      + `fill="var(--text-muted)"><title>${escapeHtml(pt.fullLabel)}</title>${escapeHtml(pt.shortLabel)}</text>`;
  }).join('');

  return `<div class="score-chart">`
    + `<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto;display:block">`
    + gridSvg
    + seriesSvg('kill_rate', 'var(--series-blue)')
    + seriesSvg('efficiency', 'var(--series-orange)')
    + xLabels
    + `</svg>`
    + `<div class="score-chart-legend">`
    + `<span class="legend-item"><span class="legend-swatch" style="background:var(--series-blue)"></span>決定率</span>`
    + `<span class="legend-item"><span class="legend-swatch" style="background:var(--series-orange)"></span>効果率</span>`
    + `</div></div>`;
}

// ==================== チーム全体の、試合をまたいだ推移グラフ ====================
// サイドアウト率・ブレイク率・アタック決定率・アタック効果率の4つを、選んだ試合を
// 日付順に並べて折れ線グラフにする。「月ごとにまとめて選ぶ」ボタンと組み合わせて使う想定。
const TEAM_TREND_SERIES = [
  { key: 'kill_rate', label: '決定率', colorVar: 'var(--series-blue)', get: m => m.team.total.kill_rate },
  { key: 'efficiency', label: '効果率', colorVar: 'var(--series-orange)', get: m => m.team.total.efficiency },
  { key: 'side_out_rate', label: 'サイドアウト率', colorVar: 'var(--series-aqua)', get: m => m.sideOutBreak.total.side_out_rate },
  { key: 'break_rate', label: 'ブレイク率', colorVar: 'var(--series-yellow)', get: m => m.sideOutBreak.total.break_rate },
];

function teamTrendChartSvg(matches) {
  const sorted = matches.slice().sort((a, b) => dvDateSlug(a.matchDate).localeCompare(dvDateSlug(b.matchDate)));
  const points = sorted.map(m => {
    const displayDate = formatMatchDate(m.matchDate);
    const parts = displayDate.split('/');
    const shortLabel = parts.length === 3 ? `${parts[1]}/${parts[2]}` : displayDate;
    const pt = { fullLabel: `${displayDate} 対 ${m.opponent}`, shortLabel };
    TEAM_TREND_SERIES.forEach(s => { pt[s.key] = s.get(m); });
    return pt;
  });
  const values = [];
  points.forEach(pt => TEAM_TREND_SERIES.forEach(s => {
    const v = pt[s.key];
    if (v !== null && v !== undefined) values.push(v);
  }));
  if (values.length < 2) return '<p class="hint">推移グラフを表示するには、試合が2試合以上必要です。</p>';

  const W = 640, H = 240;
  const padL = 42, padR = 16, padT = 14, padB = 30;
  const plotW = W - padL - padR, plotH = H - padT - padB;
  const n = points.length;

  const rawMin = Math.min(0, ...values);
  const rawMax = Math.max(0, ...values);
  const span = Math.max(0.1, rawMax - rawMin);
  const yMin = rawMin - span * 0.15;
  const yMax = rawMax + span * 0.15;

  const xAt = i => padL + (n <= 1 ? plotW / 2 : (i / (n - 1)) * plotW);
  const yAt = v => padT + plotH - ((v - yMin) / (yMax - yMin)) * plotH;

  let gridSvg = `<line x1="${padL}" y1="${padT}" x2="${padL}" y2="${(H - padB).toFixed(1)}" stroke="var(--grid)" stroke-width="1"/>`;
  if (yMin < 0 && yMax > 0) {
    const zeroY = yAt(0).toFixed(1);
    gridSvg += `<line x1="${padL}" y1="${zeroY}" x2="${W - padR}" y2="${zeroY}" stroke="var(--grid)" stroke-width="1"/>`;
  }
  gridSvg += `<text x="${(padL - 6).toFixed(1)}" y="${(yAt(yMax) + 4).toFixed(1)}" text-anchor="end" font-size="11" fill="var(--text-muted)">${pct(yMax)}</text>`;
  gridSvg += `<text x="${(padL - 6).toFixed(1)}" y="${(yAt(yMin) + 4).toFixed(1)}" text-anchor="end" font-size="11" fill="var(--text-muted)">${pct(yMin)}</text>`;

  function seriesSvg(key, colorVar) {
    let segHtml = '';
    let seg = [];
    points.forEach((pt, i) => {
      const v = pt[key];
      if (v === null || v === undefined) {
        if (seg.length > 1) {
          segHtml += `<polyline points="${seg.join(' ')}" fill="none" stroke="${colorVar}" `
            + `stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
        }
        seg = [];
        return;
      }
      seg.push(`${xAt(i).toFixed(1)},${yAt(v).toFixed(1)}`);
    });
    if (seg.length > 1) {
      segHtml += `<polyline points="${seg.join(' ')}" fill="none" stroke="${colorVar}" `
        + `stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
    }
    const dots = points.map((pt, i) => {
      const v = pt[key];
      if (v === null || v === undefined) return '';
      return `<circle cx="${xAt(i).toFixed(1)}" cy="${yAt(v).toFixed(1)}" r="3.5" fill="${colorVar}" `
        + `stroke="var(--surface-1)" stroke-width="1"><title>${escapeHtml(pt.fullLabel)}: ${pct(v)}</title></circle>`;
    }).join('');
    return segHtml + dots;
  }

  const xLabels = points.map((pt, i) => {
    const anchor = i === 0 ? 'start' : (i === n - 1 ? 'end' : 'middle');
    return `<text x="${xAt(i).toFixed(1)}" y="${H - 8}" text-anchor="${anchor}" font-size="10.5" `
      + `fill="var(--text-muted)"><title>${escapeHtml(pt.fullLabel)}</title>${escapeHtml(pt.shortLabel)}</text>`;
  }).join('');

  const seriesSvgHtml = TEAM_TREND_SERIES.map(s => seriesSvg(s.key, s.colorVar)).join('');
  const legendHtml = TEAM_TREND_SERIES.map(s =>
    `<span class="legend-item"><span class="legend-swatch" style="background:${s.colorVar}"></span>${s.label}</span>`
  ).join('');

  return `<div class="score-chart">`
    + `<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto;display:block">`
    + gridSvg + seriesSvgHtml + xLabels
    + `</svg>`
    + `<div class="score-chart-legend">${legendHtml}</div>`
    + `</div>`
    + teamTrendTableHtml(sorted);
}

// 折れ線だけだと黄色・水色が薄い色で見づらいことがあるので、正確な数値を表で併記する
function teamTrendTableHtml(sortedMatches) {
  let html = '<div class="table-wrap" style="margin-top:12px"><table><thead><tr>'
    + '<th>試合</th>' + TEAM_TREND_SERIES.map(s => `<th>${s.label}</th>`).join('') + '</tr></thead><tbody>';
  sortedMatches.forEach(m => {
    html += `<tr><td>${formatMatchDate(m.matchDate)} 対 ${escapeHtml(m.opponent)}</td>`
      + TEAM_TREND_SERIES.map(s => `<td>${pct(s.get(m))}</td>`).join('') + '</tr>';
  });
  html += '</tbody></table></div>';
  return html;
}

function playerAcrossMatchesTableHtml(number, matches) {
  const colspan = matches.length + 1;
  const statsList = matches.map(m => (m.total || []).find(pp => pp.number === number) || null);

  function row(label, getter, isPct) {
    let html = `<tr><td>${escapeHtml(label)}</td>`;
    statsList.forEach(p => {
      if (!p) { html += '<td class="hint">出場なし</td>'; return; }
      const v = getter(p);
      html += `<td>${isPct ? pct(v) : v}</td>`;
    });
    html += '</tr>';
    return html;
  }

  const hasBlock = statsList.some(p => p && p.block_attempts > 0);
  const hasServe = statsList.some(p => p && p.serve_attempts > 0);
  const hasReceive = statsList.some(p => p && p.receive_attempts > 0);

  let html = '<div style="overflow-x:auto"><table class="rotation-table comparison-table">';
  html += '<thead><tr><th>項目</th>';
  matches.forEach(m => {
    html += `<th>${escapeHtml(m.opponent)}<br><span style="font-weight:400;font-size:12px">${escapeHtml(formatMatchDate(m.matchDate))}</span></th>`;
  });
  html += '</tr></thead><tbody>';

  html += matchComparisonSectionHtml('スパイク（全セット合計）', colspan);
  html += row('打数', p => p.attempts, false);
  html += row('得点', p => p.points, false);
  html += row('ミス', p => p.errors, false);
  html += row('被ブロック', p => p.blocked, false);
  html += row('決定率', p => p.kill_rate, true);
  html += row('効果率', p => p.efficiency, true);

  if (hasBlock) {
    html += matchComparisonSectionHtml('ブロック（全セット合計）', colspan);
    html += row('本数', p => p.block_attempts, false);
    html += row('得点', p => p.block_points, false);
    html += row('ワンチ', p => p.block_touch_own + p.block_touch_opp, false);
    html += row('失点', p => p.block_errors, false);
  }
  if (hasServe) {
    html += matchComparisonSectionHtml('サーブ（全セット合計）', colspan);
    html += row('打数', p => p.serve_attempts, false);
    html += row('エース', p => p.serve_aces, false);
    html += row('ミス', p => p.serve_errors, false);
    html += row('効果本数', p => p.serve_half_credit, false);
    html += row('エース率', p => p.serve_ace_rate, true);
    html += row('ミス率', p => p.serve_error_rate, true);
    html += row('効果率', p => p.serve_efficiency, true);
  }
  if (hasReceive) {
    html += matchComparisonSectionHtml('レシーブ（全セット合計）', colspan);
    html += row('本数', p => p.receive_attempts, false);
    html += row('Aパス', p => p.receive_a, false);
    html += row('Bパス', p => p.receive_b, false);
    html += row('Cパス', p => p.receive_c, false);
    html += row('Dパス', p => p.receive_d, false);
    html += row('ミス', p => p.receive_errors, false);
    html += row('返球率', p => p.receive_return_rate, true);
    html += row('A率', p => p.receive_a_rate, true);
    html += row('AB率', p => p.receive_ab_rate, true);
  }

  html += '</tbody></table></div>';
  return html;
}

// ---- 「合算」（選択した試合をすべて足し合わせた1つの成績）を計算・表示する ----

function combineStatsList(statsList) {
  const valid = statsList.filter(Boolean);
  if (!valid.length) return null;
  const sum = {
    attempts: 0, points: 0, errors: 0, blocked: 0,
    block_attempts: 0, block_points: 0, block_errors: 0, block_touch_own: 0, block_touch_opp: 0,
    serve_attempts: 0, serve_aces: 0, serve_errors: 0, serve_half_credit: 0,
    receive_attempts: 0, receive_a: 0, receive_b: 0, receive_c: 0, receive_d: 0, receive_errors: 0,
  };
  TECH_ORDER.forEach(t => {
    sum[`tech_${t}_attempts`] = 0;
    sum[`tech_${t}_points`] = 0;
    sum[`tech_${t}_errors`] = 0;
  });
  valid.forEach(s => {
    Object.keys(sum).forEach(k => { sum[k] += (s[k] || 0); });
  });
  const result = {
    ...sum,
    kill_rate: dvRate(sum.points, sum.attempts),
    efficiency: dvRate(sum.points - sum.errors - sum.blocked, sum.attempts),
    serve_ace_rate: dvRate(sum.serve_aces, sum.serve_attempts),
    serve_error_rate: dvRate(sum.serve_errors, sum.serve_attempts),
    serve_efficiency: dvRate(sum.serve_aces - sum.serve_errors + 0.5 * sum.serve_half_credit, sum.serve_attempts),
    receive_return_rate: dvRate(sum.receive_attempts - sum.receive_errors, sum.receive_attempts),
    receive_a_rate: dvRate(sum.receive_a, sum.receive_attempts),
    receive_ab_rate: dvRate(sum.receive_a + sum.receive_b, sum.receive_attempts),
  };
  TECH_ORDER.forEach(t => {
    result[`tech_${t}_kill_rate`] = dvRate(sum[`tech_${t}_points`], sum[`tech_${t}_attempts`]);
  });
  return result;
}

function combineSideOutBreak(list) {
  const valid = list.filter(Boolean);
  const sum = { serve_total: 0, serve_wins: 0, receive_total: 0, receive_wins: 0 };
  valid.forEach(s => {
    sum.serve_total += s.serve_total; sum.serve_wins += s.serve_wins;
    sum.receive_total += s.receive_total; sum.receive_wins += s.receive_wins;
  });
  return {
    ...sum,
    break_rate: dvRate(sum.serve_wins, sum.serve_total),
    side_out_rate: dvRate(sum.receive_wins, sum.receive_total),
  };
}

// ローテーション別 攻撃タイプ分布の表を、選んだ試合分すべてで足し合わせる
// （各試合の rows は「S1〜S6 ＋ 合計」の7行で並びが揃っているので、同じ位置同士を合算する）
function combineRotationRows(rowsList) {
  const valid = rowsList.filter(r => r && r.length);
  if (!valid.length) return [];
  const rowCount = valid[0].length;
  const rows = [];
  for (let i = 0; i < rowCount; i++) {
    const sum = { attempts: 0, points: 0, errors: 0 };
    const catSum = {};
    CATEGORY_ORDER.forEach(cat => { catSum[cat] = { attempts: 0, points: 0, errors: 0 }; });
    const techSum = {};
    TECH_ORDER.forEach(t => { techSum[t] = { attempts: 0, points: 0, errors: 0 }; });
    valid.forEach(rowsForMatch => {
      const r = rowsForMatch[i];
      if (!r) return;
      sum.attempts += r.attempts; sum.points += r.points; sum.errors += r.errors;
      CATEGORY_ORDER.forEach(cat => {
        const c = r.categories[cat];
        catSum[cat].attempts += c.attempts; catSum[cat].points += c.points; catSum[cat].errors += c.errors;
      });
      TECH_ORDER.forEach(t => {
        const c = r.techniques[t];
        techSum[t].attempts += c.attempts; techSum[t].points += c.points; techSum[t].errors += c.errors;
      });
    });
    const row = {
      rotation: valid[0][i].rotation,
      attempts: sum.attempts, points: sum.points, errors: sum.errors,
      kill_rate: dvRate(sum.points, sum.attempts),
      categories: {},
    };
    CATEGORY_ORDER.forEach(cat => {
      row.categories[cat] = {
        attempts: catSum[cat].attempts, points: catSum[cat].points, errors: catSum[cat].errors,
        kill_rate: dvRate(catSum[cat].points, catSum[cat].attempts),
      };
    });
    row.techniques = {};
    TECH_ORDER.forEach(t => {
      row.techniques[t] = {
        attempts: techSum[t].attempts, points: techSum[t].points, errors: techSum[t].errors,
        kill_rate: dvRate(techSum[t].points, techSum[t].attempts),
      };
    });
    rows.push(row);
  }
  return rows;
}

// 選手個人の「ローテーション別のアタック出場状況」の表
function playerRotationTableHtml(rows) {
  const withAttempts = (rows || []).filter(r => r.attempts > 0);
  if (!withAttempts.length) return '<p class="hint">この選手のスパイク出場記録がありません。</p>';
  let html = '<div style="overflow-x:auto"><table class="rotation-table"><thead><tr>'
    + '<th>ローテ</th><th>打数</th><th>得点</th><th>ミス</th><th>決定率</th>'
    + '</tr></thead><tbody>';
  rows.forEach(r => {
    html += `<tr><td>S${r.rotation}</td><td>${r.attempts}</td><td>${r.points}</td><td>${r.errors}</td><td>${pct(r.kill_rate)}</td></tr>`;
  });
  html += '</tbody></table></div>';
  return html;
}

// 選手個人の「ローテーション別のアタック出場状況」を、選んだ試合分すべてで足し合わせる
function combinePlayerRotationRows(rowsList) {
  const valid = rowsList.filter(r => r && r.length);
  if (!valid.length) return [];
  const rows = [];
  for (let i = 0; i < 6; i++) {
    const sum = { attempts: 0, points: 0, errors: 0 };
    valid.forEach(rowsForMatch => {
      const r = rowsForMatch[i];
      if (!r) return;
      sum.attempts += r.attempts; sum.points += r.points; sum.errors += r.errors;
    });
    rows.push({
      rotation: valid[0][i].rotation,
      attempts: sum.attempts, points: sum.points, errors: sum.errors,
      kill_rate: dvRate(sum.points, sum.attempts),
    });
  }
  return rows;
}

// ローテーション別のサイドアウト率・ブレイク率を、選んだ試合分すべてで足し合わせる
function combineSideOutBreakByRotation(byRotationList) {
  const valid = byRotationList.filter(r => r && r.length);
  if (!valid.length) return [];
  const rows = [];
  for (let n = 1; n <= 6; n++) {
    const entries = valid.map(rowsForMatch => rowsForMatch.find(r => r.rotation === n)).filter(Boolean);
    rows.push({ rotation: n, ...combineSideOutBreak(entries) });
  }
  return rows;
}

// 選手ランキングを、選んだ試合分すべての選手成績の合計で作り直す
function combinedPlayerTotalsList(dataList) {
  const numbers = new Set();
  dataList.forEach(d => (d.total || []).forEach(p => numbers.add(p.number)));
  const list = [];
  numbers.forEach(number => {
    const perMatchStats = dataList.map(d => (d.total || []).find(p => p.number === number) || null);
    const combined = combineStatsList(perMatchStats);
    if (!combined) return;
    list.push({ ...combined, number, name: ROSTER[number] || `#${number}` });
  });
  return list;
}

function combinedTableHtml(stats) {
  const colspan = 2;
  let html = '<div style="overflow-x:auto"><table class="rotation-table comparison-table" style="min-width:0">';
  html += '<tbody>';

  function row(label, value, isPct) {
    return `<tr><td>${escapeHtml(label)}</td><td>${isPct ? pct(value) : value}</td></tr>`;
  }

  html += matchComparisonSectionHtml('スパイク（全セット合計）', colspan);
  html += row('打数', stats.attempts, false);
  html += row('得点', stats.points, false);
  html += row('ミス', stats.errors, false);
  html += row('被ブロック', stats.blocked, false);
  html += row('決定率', stats.kill_rate, true);
  html += row('効果率', stats.efficiency, true);

  if (stats.block_attempts > 0) {
    html += matchComparisonSectionHtml('ブロック（全セット合計）', colspan);
    html += row('本数', stats.block_attempts, false);
    html += row('得点', stats.block_points, false);
    html += row('ワンチ', stats.block_touch_own + stats.block_touch_opp, false);
    html += row('失点', stats.block_errors, false);
  }
  if (stats.serve_attempts > 0) {
    html += matchComparisonSectionHtml('サーブ（全セット合計）', colspan);
    html += row('打数', stats.serve_attempts, false);
    html += row('エース', stats.serve_aces, false);
    html += row('ミス', stats.serve_errors, false);
    html += row('効果本数', stats.serve_half_credit, false);
    html += row('エース率', stats.serve_ace_rate, true);
    html += row('ミス率', stats.serve_error_rate, true);
    html += row('効果率', stats.serve_efficiency, true);
  }
  if (stats.receive_attempts > 0) {
    html += matchComparisonSectionHtml('レシーブ（全セット合計）', colspan);
    html += row('本数', stats.receive_attempts, false);
    html += row('Aパス', stats.receive_a, false);
    html += row('Bパス', stats.receive_b, false);
    html += row('Cパス', stats.receive_c, false);
    html += row('Dパス', stats.receive_d, false);
    html += row('ミス', stats.receive_errors, false);
    html += row('返球率', stats.receive_return_rate, true);
    html += row('A率', stats.receive_a_rate, true);
    html += row('AB率', stats.receive_ab_rate, true);
  }

  html += '</tbody></table></div>';
  return html;
}

function combinedSideOutBreakHtml(s) {
  let html = '<div style="overflow-x:auto"><table class="rotation-table comparison-table" style="min-width:0"><tbody>';
  html += matchComparisonSectionHtml('サイドアウト率・ブレイク率（合計）', 2);
  html += `<tr><td>サーブ本数</td><td>${s.serve_total}</td></tr>`;
  html += `<tr><td>ブレイク率</td><td>${pct(s.break_rate)}</td></tr>`;
  html += `<tr><td>レシーブ本数</td><td>${s.receive_total}</td></tr>`;
  html += `<tr><td>サイドアウト率</td><td>${pct(s.side_out_rate)}</td></tr>`;
  html += '</tbody></table></div>';
  return html;
}

let matchComparisonOpenSections = new Set(['team-trend']);

function toggleAccordionSection(id) {
  if (matchComparisonOpenSections.has(id)) {
    matchComparisonOpenSections.delete(id);
  } else {
    matchComparisonOpenSections.add(id);
  }
  refreshMatchComparisonIfNeeded();
}

// ==================== 相手チーム分析（スカウティング） ====================
// 選んだ試合が全部同じ対戦相手の時だけ使う。次また対戦する時にどうすれば勝てそうか、
// 相手のローテーション別の強い/弱いところと、相手のよく決めている選手を見えるようにする。
function computeOpponentScorerLeaderboard(bySetList) {
  const byNumber = {};
  (bySetList || []).forEach(s => {
    (s.points || []).forEach(pt => {
      if (pt.scoringTeam !== 'opponent' || pt.byError) return;
      const n = pt.opponentScorerNumber;
      if (n === null || n === undefined) return;
      if (!byNumber[n]) byNumber[n] = { number: n, kill: 0, ace: 0, block: 0, total: 0 };
      const entry = byNumber[n];
      if (pt.skill === 'A') entry.kill += 1;
      else if (pt.skill === 'S') entry.ace += 1;
      else if (pt.skill === 'B') entry.block += 1;
      entry.total += 1;
    });
  });
  return Object.values(byNumber).sort((a, b) => b.total - a.total);
}

function opponentScorerLeaderboardHtml(list) {
  if (!list.length) return '<p class="hint">相手選手ごとの決定データがありません。</p>';
  let html = '<p class="hint" style="margin-top:0">背番号のみ（相手チームの名簿は登録していないため）。多い順。</p>'
    + '<div style="overflow-x:auto"><table class="rotation-table"><thead><tr>'
    + '<th>相手の背番号</th><th>スパイク決定</th><th>サーブエース</th><th>ブロック得点</th><th>合計</th>'
    + '</tr></thead><tbody>';
  list.slice(0, 12).forEach(e => {
    html += `<tr><td>#${e.number}</td><td>${e.kill}</td><td>${e.ace}</td><td>${e.block}</td><td><strong>${e.total}</strong></td></tr>`;
  });
  html += '</tbody></table></div>';
  return html;
}

// 相手のローテーション別サイドアウト率・ブレイク率から、狙い目を一文にする
function opponentRotationInsightText(rows) {
  const MIN = 3;
  const parts = [];
  const soRows = rows.filter(r => r.receive_total >= MIN && r.side_out_rate !== null && r.side_out_rate !== undefined);
  if (soRows.length >= 2) {
    const worst = soRows.reduce((a, b) => (b.side_out_rate < a.side_out_rate ? b : a));
    parts.push(`相手はS${worst.rotation}で受けている時が一番サイドアウトしにくい(${pct(worst.side_out_rate)})ので、そこを狙ってサーブを集めるのが効果的かもしれません`);
  }
  const brRows = rows.filter(r => r.serve_total >= MIN && r.break_rate !== null && r.break_rate !== undefined);
  if (brRows.length >= 2) {
    const best = brRows.reduce((a, b) => (b.break_rate > a.break_rate ? b : a));
    parts.push(`相手はS${best.rotation}で自分たちがサーブの時に一番ブレイクしやすい(${pct(best.break_rate)})ので、このローテーションは特に丁寧に返したいところです`);
  }
  return parts.length ? '注目ポイント： ' + parts.join('／') : '';
}

function opponentAnalysisHtml(matches) {
  const combinedByRotation = combineSideOutBreakByRotation(matches.map(m => m.opponentSideOutBreak && m.opponentSideOutBreak.byRotation || []));
  const insight = opponentRotationInsightText(combinedByRotation);
  const bySetForPattern = matches.flatMap(m => m.scoreProgression.bySet);

  let html = '<h4 class="section-label" style="margin-top:0;font-size:15px">相手のローテーション別成績</h4>';
  if (insight) html += `<p class="hint" style="margin-top:-4px">${escapeHtml(insight)}</p>`;
  html += sideOutBreakByRotationHtml(combinedByRotation);
  // 「相手のよく決めている選手」（背番号ランキング）は2026-08-21にいったん削除。
  // VolleyStationの相手チーム名簿が対戦相手によらず毎回同じ仮の名簿（YAMADA/SATOU...の1〜7番）の
  // ままになっており、本当の背番号が入力されていないため、複数試合をまたいで背番号を集計しても
  // 同じ選手とは限らず、信頼できるデータではなかった（こうせいさんの指摘で判明）。
  // スカウティング時に相手チームの本当の背番号を入力してもらえるようになったら、
  // computeOpponentScorerLeaderboard / opponentScorerLeaderboardHtml はそのまま使えるので復活できる。
  html += '<h4 class="section-label" style="font-size:15px">得点パターン（この相手との対戦分）</h4>';
  html += scoringPatternHtml(computeScoringPattern(bySetForPattern));
  return html;
}

let matchComparisonAllSectionIds = [];

function accordionSectionHtml(id, title, bodyHtml) {
  const isOpen = matchComparisonOpenSections.has(id);
  if (!matchComparisonAllSectionIds.includes(id)) matchComparisonAllSectionIds.push(id);
  return `
    <div class="accordion-section">
      <button type="button" class="accordion-toggle" data-section-id="${id}" onclick="toggleAccordionSection('${id}')">
        <span>${escapeHtml(title)}</span><span class="arrow">${isOpen ? '▴' : '▾'}</span>
      </button>
      <div class="accordion-body" style="display:${isOpen ? 'block' : 'none'}">${bodyHtml}</div>
    </div>`;
}

function toggleAllAccordionSections(btnEl) {
  const willExpand = btnEl.textContent.trim() !== 'すべて折りたたむ';
  if (willExpand) {
    matchComparisonAllSectionIds.forEach(id => matchComparisonOpenSections.add(id));
  } else {
    matchComparisonOpenSections.clear();
  }
  btnEl.textContent = willExpand ? 'すべて折りたたむ' : 'すべて展開';
  refreshMatchComparisonIfNeeded();
}

function showMatchComparison(matches) {
  document.getElementById('matchComparisonCard').style.display = 'block';
  document.getElementById('matchComparisonTitle').textContent = `試合比較（${matches.length}試合）`;

  let html = accordionSectionHtml('team-trend', 'チームの推移（サイドアウト率・ブレイク率・決定率・効果率）', teamTrendChartSvg(matches));

  // 相手チーム分析（スカウティング）セクションは2026-08-29にいったん非表示に。
  // opponentAnalysisHtml関数自体はそのまま残っているので、また使いたくなったら
  // 下の3行のコメントを外せば復活できる。
  // const opponentNames = new Set(matches.map(m => m.opponent));
  // if (opponentNames.size === 1) {
  //   const oppName = matches[0].opponent;
  //   html += accordionSectionHtml('opponent-analysis', `対 ${oppName} 戦の傾向（スカウティング）`, opponentAnalysisHtml(matches));
  // }

  html += accordionSectionHtml('team-compare', '試合ごとの比較（チーム全体）', matchComparisonTableHtml(matches));

  const combinedTeam = combineStatsList(matches.map(m => m.team.total));
  const combinedSOB = combineSideOutBreak(matches.map(m => m.sideOutBreak.total));
  let combinedHtml = '';
  if (combinedTeam) combinedHtml += combinedTableHtml(combinedTeam);
  combinedHtml += combinedSideOutBreakHtml(combinedSOB);
  html += accordionSectionHtml('team-combined', `合算スタッツ（選んだ${matches.length}試合をすべて合計・チーム全体）`, combinedHtml);

  Array.from(selectedPlayerNumbers).forEach(number => {
    const name = ROSTER[number] || `#${number}`;
    let playerHtml = '<h4 class="section-label" style="margin-top:0;font-size:15px">試合をまたいだ推移（決定率・効果率）</h4>';
    playerHtml += playerTrendChartSvg(number, matches);
    playerHtml += '<h4 class="section-label" style="font-size:15px">試合ごとの比較</h4>';
    playerHtml += playerAcrossMatchesTableHtml(number, matches);
    const playerStatsList = matches.map(m => (m.total || []).find(pp => pp.number === number) || null);
    const combinedPlayer = combineStatsList(playerStatsList);
    playerHtml += `<h4 class="section-label" style="font-size:15px">合算スタッツ（選んだ${matches.length}試合をすべて合計）</h4>`;
    playerHtml += combinedPlayer ? combinedTableHtml(combinedPlayer) : '<p class="hint">選んだ試合に出場記録がありません</p>';
    html += accordionSectionHtml(`player-${number}`, `${name}（#${number}）`, playerHtml);
  });

  document.getElementById('matchComparisonBody').innerHTML = html;
}

document.getElementById('teamTitle').textContent = '__TEAM_LABEL__';
setupDropZone();
setupDropZoneToggle();
setupPlayerListToggle();
setupSidebarToggle();
setupMatchListToggle();
// 対戦相手クイック選択（スカウティング用）は2026-08-29にいったん非表示に。
// setupOpponentQuickSelect();

fetch('matches.json')
  .then(res => res.ok ? res.json() : [])
  .catch(() => [])
  .then(matches => {
    MATCHES = matches || [];
    if (!matches || !matches.length) {
      document.getElementById('noMatchesHint').style.display = 'block';
      return;
    }
    selectedMatchKeys.add(matchKey(matches[0]));
    renderMatchChecklist();
    // renderOpponentQuickSelect();
    renderMatchSelection();
  });
</script>
</body>
</html>
"""


def load_file(path):
    """.dvwファイルを読み込んで、中身の文字列を返す"""
    with open(path, encoding='ascii') as f:
        return f.read()


def get_section(content, name):
    """[3XXXX] というセクションの中身だけを取り出す"""
    start = content.find(f'[{name}]')
    if start == -1:
        return ''
    start += len(f'[{name}]\n')
    end = content.find('[3', start)
    if end == -1:
        end = len(content)
    return content[start:end]


# 自分たちのチームコード。.dvwファイルの[3TEAMS]セクションに出てくるコードで、
# 女子は'TSJ'、男子は'TSD'（男子の試合データで確認済み）。チームごとに違うので
# TEAMの設定に応じて自動的に切り替わるようにしてある
TARGET_TEAM_CODE_BY_TEAM = {'women': 'TSJ', 'men': 'TSD'}
TARGET_TEAM_CODE = TARGET_TEAM_CODE_BY_TEAM[TEAM]


def get_own_team_marker(content):
    """
    [3TEAMS]セクションを見て、自チーム(TARGET_TEAM_CODE)が
    ホーム('*')なのかアウェイ('a')なのかを調べる
    """
    teams_section = get_section(content, '3TEAMS')
    team_lines = [l for l in teams_section.splitlines() if l.strip()]
    team_codes = [l.split(';')[0] for l in team_lines]
    if team_codes[0] == TARGET_TEAM_CODE:
        return '*'
    else:
        return 'a'


def get_back_attack_combos(content):
    """
    [3ATTACKCOMBINATION]から、
    「バックアタック(後衛からの攻撃)」に該当する攻撃コードの集合を作る。

    判定ルール：コンボ表の2番目のフィールドが「ゾーン番号」で、
    7・8・9番はコートの後方(バックアタック用)ゾーンを表す。
    （最初は9番目のフィールド=カテゴリ文字'P'で判定していたが、
    実データの答え合わせをしたところ、P9やL9などゾーン9のコードが
    'B'カテゴリになっていて漏れることが判明。ゾーン番号での判定に修正）
    """
    section = get_section(content, '3ATTACKCOMBINATION')
    back_combos = set()
    for line in section.splitlines():
        if not line.strip():
            continue
        fields = line.split(';')
        combo_code = fields[0]
        zone = fields[1] if len(fields) > 1 else ''
        if zone in ('7', '8', '9'):
            back_combos.add(combo_code)
    return back_combos


def split_into_sets(lines):
    """
    '**1set' '**2set' のような区切り行を目印に、
    セットごとの行リストに分割する
    """
    sets = []
    current = []
    for line in lines:
        # 行全体を残すようにしたので、区切り行の後ろに得点情報などが
        # ついていても検知できるよう、末尾の $ 完全一致は外している
        if re.match(r'^\*\*\dset', line):
            sets.append(current)
            current = []
        else:
            current.append(line)
    if current:
        sets.append(current)
    return sets


# アタック(A)・ブロック(B)・サーブ(S)・レシーブ(R) の行にマッチする正規表現
# 例: *07AH#P3~3~~H2  → team=*, player=07, skill=A, type=H, eval=#, combo=P3
# 例: *07BM!~~~~2B~2  → team=*, player=07, skill=B, type=M, eval=!
# 例: *06SM-          → team=*, player=06, skill=S, type=M, eval=-
# 例: *04RM=          → team=*, player=04, skill=R, type=M, eval==
ACTION_RE = re.compile(r'^([\*a])(\d\d)([ABSR])([A-Za-z])([#+\-!=/])([A-Z0-9]{2})?')

# 得点パターン集計（analyze_score_progression）専用：ディグ(D)・セット(E)・フリーボール(F)も
# 含めた「全スキル」にマッチする正規表現。得点/失点の直前にどのプレーがあったかを追うには、
# アタック・ブロック・サーブ・レシーブだけでなく、ディグミスなどでラリーが終わるケースも
# 取りこぼさずに拾う必要があるため、ACTION_REとは別に用意している。
ACTION_ANY_SKILL_RE = re.compile(r'^([\*a])(\d\d)([A-Z])([A-Za-z])([#+\-!=/])([A-Z0-9]{2})?')

# 得点が入った瞬間の行 例: '*p08:05' → ホームが得点して8-5に、'ap03:10' → アウェイが得点して3-10に
SCORE_RE = re.compile(r'^([\*a])p(\d+):(\d+)')

# 選手1人分の集計の初期値（この形をずっと使い回す）
EMPTY_PLAYER_STATS = {
    'attempts': 0, 'points': 0, 'errors': 0, 'blocked': 0,
    'back_attempts': 0, 'back_points': 0, 'back_errors': 0,
    'block_attempts': 0, 'block_points': 0, 'block_errors': 0,
    'block_touch_own': 0, 'block_touch_opp': 0,
    'serve_attempts': 0, 'serve_aces': 0, 'serve_errors': 0, 'serve_half_credit': 0,
    'receive_attempts': 0, 'receive_a': 0, 'receive_b': 0,
    'receive_c': 0, 'receive_d': 0, 'receive_errors': 0,
    # 打法（強打／フェイント／その他）別の選手ごとの内訳（2026-08-29追加）
    'tech_hard_attempts': 0, 'tech_hard_points': 0, 'tech_hard_errors': 0,
    'tech_feint_attempts': 0, 'tech_feint_points': 0, 'tech_feint_errors': 0,
    'tech_other_attempts': 0, 'tech_other_points': 0, 'tech_other_errors': 0,
}

# 攻撃コンビ表([3ATTACKCOMBINATION])の9番目の項目(カテゴリ文字)を、
# レフト／ライト／ミドル／バックアタック／その他 に振り分ける対応表
# 実データで突き合わせた結果：F=レフト系(4番・7番などコートの左サイド)
# B=ライト系(2番・9番などコートの右サイド) C=ミドル/クイック系 P=バックアタック(内部キーはpipeのまま)
# それ以外(セッター自身の攻撃など、ごく少数)は「その他」にまとめている
CATEGORY_ORDER = ['left', 'right', 'middle', 'pipe', 'other']
CATEGORY_LABELS = {'left': 'レフト', 'right': 'ライト', 'middle': 'ミドル', 'pipe': 'バックアタック', 'other': 'その他'}
_COMBO_CATEGORY_MAP = {'F': 'left', 'B': 'right', 'C': 'middle', 'P': 'pipe'}
# [3ATTACKCOMBINATION]は同じコンビコードが複数回定義されていることがある
# （VolleyStationのテンプレート由来。例：'P2'は「短いレフトのハイセット」＝通常のレフト攻撃
# なのだが、テンプレート内にもう一つ紛らわしい定義（緊急時2段トスの説明文）が重複して
# 入っている）。コード→カテゴリの対応表は最後に出てきた定義を採用する単純な仕組みのままにし、
# 紛らわしいものだけ以下で明示的に上書きする（2026-08-29、こうせいさんに確認済み）。
# P9・L9はコンビ表では分類文字'B'（ライト）扱いだが、実際はゾーン9＝後衛からのバックライト攻撃
# なので「バックアタック」（内部キーは従来のpipeのまま、表示名だけ変更）にまとめる。
# P3は分類文字が空欄だが実際にはミドル攻撃として使われている（いずれもこうせいさんに確認、2026-08-31）。
COMBO_CATEGORY_OVERRIDES = {'P2': 'left', 'P9': 'pipe', 'L9': 'pipe', 'P3': 'middle'}

# 打法（強打／フェイント）の分類。H=強打、T=フェイントであることをこうせいさんに確認（2026-08-29）。
# ※このH/Tは、スキル直後の1文字(例 "*05AH#P5"の"H")ではなく、コードの末尾側
# （'~'区切りの最後のかたまり。例 "*05AH#P5~24~H2"の"H2"、"a02AM+PV~4~~T2"の"T2"）に
# 入っている、実際の「ショットタイプ」の文字。スキル直後の文字は実データを見るとM/H/Q/N/Oで、
# こちらは(H/M/Q/N/Oの)「トスのテンポ」を表しており、強打/フェイントの区別ではなかった
# （2026-08-29、実データで確認）。この2つ＋その他、で攻撃の打法を分類する。
TECH_ORDER = ['hard', 'feint', 'other']
TECH_LABELS = {'hard': '強打', 'feint': 'フェイント', 'other': 'その他'}
_ATTACK_SHOT_TYPE_RE = re.compile(r'([A-Z])(\d)')


def attack_shot_type(code):
    """アタックの生コード全体（例 '*05AH#P5~24~H2'）から、末尾の「ショットタイプ」文字を取り出す。
    '~'で区切った最後のかたまり（ゾーン番号や区分け文字が前についていることがある。
    例：'H2'、'42BH2'、'24CH2'、'T2N'など）の中から、数字が直後に続く最初のアルファベットを探す。
    """
    parts = code.split('~')
    if len(parts) < 2:
        return None
    m = _ATTACK_SHOT_TYPE_RE.search(parts[-1])
    return m.group(1) if m else None


def get_combo_categories(content):
    """[3ATTACKCOMBINATION]セクションから、コンビコード→攻撃タイプ の対応表を作る。"""
    section = get_section(content, '3ATTACKCOMBINATION')
    mapping = {}
    for line in section.splitlines():
        if not line.strip():
            continue
        fields = line.split(';')
        code = fields[0]
        category_letter = fields[8] if len(fields) > 8 else ''
        mapping[code] = _COMBO_CATEGORY_MAP.get(category_letter, 'other')
    mapping.update(COMBO_CATEGORY_OVERRIDES)
    return mapping


def empty_rotation_entry():
    return {
        'attempts': 0, 'points': 0, 'errors': 0,
        'categories': {cat: {'attempts': 0, 'points': 0, 'errors': 0} for cat in CATEGORY_ORDER},
        'techniques': {t: {'attempts': 0, 'points': 0, 'errors': 0} for t in TECH_ORDER},
    }


def analyze_set(lines_in_set, own_marker, back_attack_combos):
    """1セット分の行から、選手ごとのアタック・ブロック集計をする"""
    # 選手ごとの集計を入れる辞書。まだ出てきていない選手はここで初期化される
    stats = {}

    def get_player_stats(number):
        if number not in stats:
            stats[number] = dict(EMPTY_PLAYER_STATS)
        return stats[number]

    for line in lines_in_set:
        fields = line.split(';')
        code = fields[0]
        m = ACTION_RE.match(code)
        if not m:
            continue
        team, player_num, skill, skill_type, evaluation, combo = m.groups()

        # ホームチーム('*')かどうかで、スリジエの選手かを判定
        # ※away('a')側がスリジエの試合では、下の '==' を '!=' に直してください
        if team != own_marker:
            continue

        number = int(player_num)
        p = get_player_stats(number)

        if skill == 'A':
            is_back = combo in back_attack_combos
            p['attempts'] += 1
            if is_back:
                p['back_attempts'] += 1
            if evaluation == '#':
                p['points'] += 1
                if is_back:
                    p['back_points'] += 1
            elif evaluation == '=':
                p['errors'] += 1
                if is_back:
                    p['back_errors'] += 1
            elif evaluation == '/':
                p['blocked'] += 1

            # 打法（強打／フェイント／その他）別の内訳も選手ごとに集計する
            shot = attack_shot_type(code)
            tech = 'hard' if shot == 'H' else ('feint' if shot == 'T' else 'other')
            p[f'tech_{tech}_attempts'] += 1
            if evaluation == '#':
                p[f'tech_{tech}_points'] += 1
            elif evaluation == '=':
                p[f'tech_{tech}_errors'] += 1

        elif skill == 'B':
            # ブロックの判定ルール（実データとスコア推移の答え合わせ済み）：
            # '#'=ブロック得点  '='=ブロック失点
            # '!'=ワンチ・相手コートへ  '+'=ワンチ・自コートへ
            # '/'も失点として扱う（こうせいさんの指定、2026-08-29〜の試合から）
            p['block_attempts'] += 1
            if evaluation == '#':
                p['block_points'] += 1
            elif evaluation == '=':
                p['block_errors'] += 1
            elif evaluation == '/':
                p['block_errors'] += 1
            elif evaluation == '!':
                p['block_touch_opp'] += 1
            elif evaluation == '+':
                p['block_touch_own'] += 1

        elif skill == 'S':
            # サーブの判定ルール（得点推移で答え合わせ済み）：
            # '#'=サーブエース  '='=サーブミス
            # '+'・'/' は、エースではないが相手のレシーブを崩せている効果的なサーブなので、
            # 効果率にはエースの半分の重みで加える。'-'は相手にAパスされている（＝崩せていない）ので対象外
            # （こうせいさんの指定、2026-08-30。以前は'-'を含めていたが変更）
            p['serve_attempts'] += 1
            if evaluation == '#':
                p['serve_aces'] += 1
            elif evaluation == '=':
                p['serve_errors'] += 1
            elif evaluation in ('+', '/'):
                p['serve_half_credit'] += 1

        elif skill == 'R':
            # レシーブ（サーブレシーブ）の判定ルール（得点推移で答え合わせ済み）：
            # 相手のサーブが '#'(エース) のときは、必ず直後に自チームの
            # レシーブが '=' として記録される（＝返せなかった、が自動で対になっている）
            # ので、レシーブの'='は「返球できずそのまま失点」で間違いない。
            # ここから先はご指定のパス評価に合わせた表記：
            # '#'=Aパス（パーフェクト） '+'=Bパス（グッド） '!'=Cパス
            # '-'と'/'=Dパス
            # （2026-08-29〜の試合から評価基準を変更。こうせいさんの指定）
            p['receive_attempts'] += 1
            if evaluation == '#':
                p['receive_a'] += 1
            elif evaluation == '+':
                p['receive_b'] += 1
            elif evaluation == '=':
                p['receive_errors'] += 1
            elif evaluation == '!':
                p['receive_c'] += 1
            elif evaluation in ('-', '/'):
                p['receive_d'] += 1
            else:
                p['receive_c'] += 1

    return stats


def analyze_rotation_attacks(lines_in_set, own_marker, combo_categories):
    """
    1セット分の行から、自チームのローテーション(S1～S6)ごと・攻撃タイプごとの
    アタック集計をする。

    「ローテーション」は、セット開始時点のならびをS1として、自チームの
    ローテーション配列(行末の背番号の並び)が変わるたび(＝1回転するたび)に
    S2→S3→…→S6→S1 と数える。得点を取ったのが誰か・セッターが誰かには
    依存しない、純粋に「自チームが何回転目か」だけの数え方。
    """
    rotation_stats = {n: empty_rotation_entry() for n in range(1, 7)}
    own_rotation = None
    rotation_number = 1

    for line in lines_in_set:
        fields = line.split(';')
        code = fields[0]

        if len(fields) >= 26:
            own_fields = tuple(fields[14:20]) if own_marker == '*' else tuple(fields[20:26])
            if own_fields != ('', '', '', '', '', ''):
                if own_rotation is None:
                    own_rotation = own_fields
                elif own_fields != own_rotation:
                    own_rotation = own_fields
                    rotation_number = rotation_number % 6 + 1

        m = ACTION_RE.match(code)
        if not m:
            continue
        team, player_num, skill, skill_type, evaluation, combo = m.groups()
        if team != own_marker or skill != 'A':
            continue
        # PPはセッターダンプではなく「ダイレクト」攻撃（トスの型を経ていないアドリブ的な返球）で、
        # トス配分・打法別・ローテーション別の集計にはなじまないため、この集計からは丸ごと除外する
        # （選手個人の通常のスパイク統計にはこれまで通りカウントされる。こうせいさんの指定、2026-08-31）
        if combo == 'PP':
            continue

        entry = rotation_stats[rotation_number]
        entry['attempts'] += 1
        if evaluation == '#':
            entry['points'] += 1
        elif evaluation == '=':
            entry['errors'] += 1

        cat_entry = entry['categories'][combo_categories.get(combo, 'other')]
        cat_entry['attempts'] += 1
        if evaluation == '#':
            cat_entry['points'] += 1
        elif evaluation == '=':
            cat_entry['errors'] += 1

        shot = attack_shot_type(code)
        tech = 'hard' if shot == 'H' else ('feint' if shot == 'T' else 'other')
        tech_entry = entry['techniques'][tech]
        tech_entry['attempts'] += 1
        if evaluation == '#':
            tech_entry['points'] += 1
        elif evaluation == '=':
            tech_entry['errors'] += 1

    return rotation_stats


def merge_rotation_stats(all_set_rotation_stats):
    """複数セット分のローテーション集計を、ローテーション番号ごとに合計する"""
    total = {n: empty_rotation_entry() for n in range(1, 7)}
    for set_stats in all_set_rotation_stats:
        for n in range(1, 7):
            for key in ('attempts', 'points', 'errors'):
                total[n][key] += set_stats[n][key]
            for cat in CATEGORY_ORDER:
                for key in ('attempts', 'points', 'errors'):
                    total[n]['categories'][cat][key] += set_stats[n]['categories'][cat][key]
            for tech in TECH_ORDER:
                for key in ('attempts', 'points', 'errors'):
                    total[n]['techniques'][tech][key] += set_stats[n]['techniques'][tech][key]
    return total


def analyze_player_rotation_attacks(lines_in_set, own_marker):
    """
    1セット分の行から、自チームのローテーション(S1〜S6)ごと・選手ごとの
    アタック集計をする（誰がどのローテーションで何本打っているかを見るため）。
    ローテーションの数え方は analyze_rotation_attacks と同じ。
    """
    result = {n: {} for n in range(1, 7)}
    own_rotation = None
    rotation_number = 1

    for line in lines_in_set:
        fields = line.split(';')
        code = fields[0]

        if len(fields) >= 26:
            own_fields = tuple(fields[14:20]) if own_marker == '*' else tuple(fields[20:26])
            if own_fields != ('', '', '', '', '', ''):
                if own_rotation is None:
                    own_rotation = own_fields
                elif own_fields != own_rotation:
                    own_rotation = own_fields
                    rotation_number = rotation_number % 6 + 1

        m = ACTION_RE.match(code)
        if not m:
            continue
        team, player_num, skill, skill_type, evaluation, combo = m.groups()
        if team != own_marker or skill != 'A':
            continue

        number = int(player_num)
        players = result[rotation_number]
        if number not in players:
            players[number] = {'attempts': 0, 'points': 0, 'errors': 0}
        p = players[number]
        p['attempts'] += 1
        if evaluation == '#':
            p['points'] += 1
        elif evaluation == '=':
            p['errors'] += 1

    return result


def merge_player_rotation_attacks(all_set_results):
    """複数セット分の「選手×ローテーション」集計を合計する"""
    total = {n: {} for n in range(1, 7)}
    for set_result in all_set_results:
        for n in range(1, 7):
            for number, s in set_result[n].items():
                if number not in total[n]:
                    total[n][number] = {'attempts': 0, 'points': 0, 'errors': 0}
                t = total[n][number]
                t['attempts'] += s['attempts']
                t['points'] += s['points']
                t['errors'] += s['errors']
    return total


def build_player_rotation_summary(merged_player_rotation):
    """選手ごとに引きやすいよう { 選手番号: [6ローテーション分の行] } の形にまとめ直す"""
    numbers = set()
    for n in range(1, 7):
        numbers.update(merged_player_rotation[n].keys())
    summary = {}
    for number in numbers:
        rows = []
        for n in range(1, 7):
            s = merged_player_rotation[n].get(number, {'attempts': 0, 'points': 0, 'errors': 0})
            rows.append({
                'rotation': n,
                'attempts': s['attempts'],
                'points': s['points'],
                'errors': s['errors'],
                'kill_rate': rate(s['points'], s['attempts']),
            })
        summary[number] = rows
    return summary


def rotation_summary(rotation_stats):
    """ローテーション集計(1～6の辞書)に決定率を付け加え、表示用の行リストにする
    （末尾に6ローテーション分を合計した「合計」行も追加する）"""
    rows = []
    grand = empty_rotation_entry()
    for n in range(1, 7):
        e = rotation_stats[n]
        row = {
            'rotation': n,
            'attempts': e['attempts'],
            'points': e['points'],
            'errors': e['errors'],
            'kill_rate': rate(e['points'], e['attempts']),
            'categories': {
                cat: {
                    'attempts': e['categories'][cat]['attempts'],
                    'points': e['categories'][cat]['points'],
                    'errors': e['categories'][cat]['errors'],
                    'kill_rate': rate(e['categories'][cat]['points'], e['categories'][cat]['attempts']),
                }
                for cat in CATEGORY_ORDER
            },
            'techniques': {
                tech: {
                    'attempts': e['techniques'][tech]['attempts'],
                    'points': e['techniques'][tech]['points'],
                    'errors': e['techniques'][tech]['errors'],
                    'kill_rate': rate(e['techniques'][tech]['points'], e['techniques'][tech]['attempts']),
                }
                for tech in TECH_ORDER
            },
        }
        rows.append(row)
        grand['attempts'] += e['attempts']
        grand['points'] += e['points']
        grand['errors'] += e['errors']
        for cat in CATEGORY_ORDER:
            for key in ('attempts', 'points', 'errors'):
                grand['categories'][cat][key] += e['categories'][cat][key]
        for tech in TECH_ORDER:
            for key in ('attempts', 'points', 'errors'):
                grand['techniques'][tech][key] += e['techniques'][tech][key]

    rows.append({
        'rotation': '合計',
        'attempts': grand['attempts'],
        'points': grand['points'],
        'errors': grand['errors'],
        'kill_rate': rate(grand['points'], grand['attempts']),
        'categories': {
            cat: {
                'attempts': grand['categories'][cat]['attempts'],
                'points': grand['categories'][cat]['points'],
                'errors': grand['categories'][cat]['errors'],
                'kill_rate': rate(grand['categories'][cat]['points'], grand['categories'][cat]['attempts']),
            }
            for cat in CATEGORY_ORDER
        },
        'techniques': {
            tech: {
                'attempts': grand['techniques'][tech]['attempts'],
                'points': grand['techniques'][tech]['points'],
                'errors': grand['techniques'][tech]['errors'],
                'kill_rate': rate(grand['techniques'][tech]['points'], grand['techniques'][tech]['attempts']),
            }
            for tech in TECH_ORDER
        },
    })
    return rows


def analyze_side_out_break(lines_in_set, own_marker):
    """
    1セット分の行から、自チームが「サーブ側だったラリー」と
    「レシーブ側だったラリー」それぞれについて、本数と勝った本数を数える。

    サーブ側で勝つ＝ブレイク、レシーブ側で勝つ＝サイドアウト。
    得点が入った瞬間の行('*pXX:YY'/'apXX:YY')の直前に出てきた最新の
    サーブ('S')が、そのラリーの「どちらが打ったサーブか」を表す。
    """
    result = {'serve_total': 0, 'serve_wins': 0, 'receive_total': 0, 'receive_wins': 0}
    current_server = None

    for line in lines_in_set:
        code = line.split(';')[0]
        m_action = ACTION_RE.match(code)
        if m_action and m_action.group(3) == 'S':
            current_server = m_action.group(1)

        m_score = SCORE_RE.match(code)
        if m_score and current_server is not None:
            scorer = m_score.group(1)
            if current_server == own_marker:
                result['serve_total'] += 1
                if scorer == own_marker:
                    result['serve_wins'] += 1
            else:
                result['receive_total'] += 1
                if scorer == own_marker:
                    result['receive_wins'] += 1

    return result


def merge_side_out_break(all_set_results):
    """複数セット分のサイドアウト/ブレイク集計を合計する"""
    total = {'serve_total': 0, 'serve_wins': 0, 'receive_total': 0, 'receive_wins': 0}
    for r in all_set_results:
        for key in total:
            total[key] += r[key]
    return total


def analyze_side_out_break_by_rotation(lines_in_set, own_marker):
    """
    1セット分の行から、自チームのローテーション(S1～S6)ごとに
    サイドアウト率・ブレイク率を集計する。
    ローテーションの数え方は analyze_rotation_attacks と同じ
    （自チームの背番号の並びが変わるたびに1回転とみなす）。
    """
    result = {n: {'serve_total': 0, 'serve_wins': 0, 'receive_total': 0, 'receive_wins': 0} for n in range(1, 7)}
    own_rotation = None
    rotation_number = 1
    current_server = None

    for line in lines_in_set:
        fields = line.split(';')
        code = fields[0]

        if len(fields) >= 26:
            own_fields = tuple(fields[14:20]) if own_marker == '*' else tuple(fields[20:26])
            if own_fields != ('', '', '', '', '', ''):
                if own_rotation is None:
                    own_rotation = own_fields
                elif own_fields != own_rotation:
                    own_rotation = own_fields
                    rotation_number = rotation_number % 6 + 1

        m_action = ACTION_RE.match(code)
        if m_action and m_action.group(3) == 'S':
            current_server = m_action.group(1)

        m_score = SCORE_RE.match(code)
        if m_score and current_server is not None:
            scorer = m_score.group(1)
            entry = result[rotation_number]
            if current_server == own_marker:
                entry['serve_total'] += 1
                if scorer == own_marker:
                    entry['serve_wins'] += 1
            else:
                entry['receive_total'] += 1
                if scorer == own_marker:
                    entry['receive_wins'] += 1

    return result


def merge_side_out_break_by_rotation(all_set_results):
    """複数セット分の、ローテーション別サイドアウト/ブレイク集計をローテーション番号ごとに合計する"""
    total = {n: {'serve_total': 0, 'serve_wins': 0, 'receive_total': 0, 'receive_wins': 0} for n in range(1, 7)}
    for set_result in all_set_results:
        for n in range(1, 7):
            for key in total[n]:
                total[n][key] += set_result[n][key]
    return total


def side_out_break_summary(r):
    """サーブ本数・勝利数などから、ブレイク率・サイドアウト率を付け加える"""
    return {
        'serve_total': r['serve_total'],
        'serve_wins': r['serve_wins'],
        'break_rate': rate(r['serve_wins'], r['serve_total']),
        'receive_total': r['receive_total'],
        'receive_wins': r['receive_wins'],
        'side_out_rate': rate(r['receive_wins'], r['receive_total']),
    }


def analyze_score_progression(lines_in_set, own_marker):
    """
    1セット分の得点推移を、ラリーが決まるたびに記録していく
    （0-0から始まり、1点入るごとに1つずつ増える）。
    得点行('*pXX:YY'/'apXX:YY')のXXは常にホーム側の点数、YYは常にアウェイ側の点数
    （どちらが決めた点かに関わらずこの順番）なので、own_markerがホーム('*')か
    アウェイ('a')かで、どちらを自チームの点数として読むかを決める。

    あわせて、その得点/失点に自チームのどの選手が絡んだかも記録する。
    得点行の直前にある自チームの決定('#')ならその選手が得点者、直前にある
    自チームのミス('=')ならそのミスをした選手として扱う（相手のミスによる
    得点は、選手を特定せず「相手のミス」として扱う）。
    """
    points = [{
        'own': 0, 'opponent': 0, 'scoringTeam': None, 'scorerNumber': None, 'byError': False, 'skill': None,
        'opponentScorerNumber': None,
    }]
    last_team = last_number = last_eval = last_skill = None
    for line in lines_in_set:
        fields = line.split(';')
        code = fields[0]

        m_action = ACTION_ANY_SKILL_RE.match(code)
        if m_action:
            team, player_num, skill, skill_type, evaluation, combo = m_action.groups()
            if evaluation in ('#', '='):
                last_team, last_number, last_eval, last_skill = team, int(player_num), evaluation, skill

        m = SCORE_RE.match(code)
        if not m:
            continue
        home_score = int(m.group(2))
        away_score = int(m.group(3))
        if own_marker == '*':
            own_score, opponent_score = home_score, away_score
        else:
            own_score, opponent_score = away_score, home_score

        prev = points[-1]
        scoring_team = scorer_number = skill_used = opponent_scorer_number = None
        by_error = False
        if own_score > prev['own']:
            scoring_team = 'own'
            if last_team == own_marker and last_eval == '#':
                scorer_number = last_number
                skill_used = last_skill
            elif last_team is not None and last_team != own_marker and last_eval == '=':
                by_error = True
                skill_used = last_skill
        elif opponent_score > prev['opponent']:
            scoring_team = 'opponent'
            if last_team == own_marker and last_eval == '=':
                scorer_number = last_number
                by_error = True
                skill_used = last_skill
            elif last_team is not None and last_team != own_marker and last_eval == '#':
                skill_used = last_skill
                opponent_scorer_number = last_number

        points.append({
            'own': own_score, 'opponent': opponent_score,
            'scoringTeam': scoring_team, 'scorerNumber': scorer_number, 'byError': by_error,
            'skill': skill_used, 'opponentScorerNumber': opponent_scorer_number,
        })
        last_team = last_number = last_eval = last_skill = None
    return points


def get_starting_lineup(lines_in_set, own_marker):
    """
    そのセット開始時点の、自チームのスタメン(P1〜P6の背番号)を取り出す。
    ローテーション追跡(analyze_rotation_attacks)と同じ、行末のならび
    (fields[14:20]がホーム、fields[20:26]がアウェイ)を使う。
    そのセットで最初に6人分そろって出てくる行が、そのセットのスタメンにあたる。
    """
    for line in lines_in_set:
        fields = line.split(';')
        if len(fields) < 26:
            continue
        own_fields = fields[14:20] if own_marker == '*' else fields[20:26]
        if any(v == '' for v in own_fields):
            continue
        return [int(v) for v in own_fields]
    return None


def parse_wallclock_seconds(t):
    """'10.05.28' のような時刻表記を、真夜中からの経過秒数に変換する"""
    h, m, s = t.split('.')
    return int(h) * 3600 + int(m) * 60 + int(s)


def build_play_log(lines_in_set, own_marker):
    """
    映像連携（プレーを押すとその場面の動画に飛べる機能）用に、自チームの
    アタック・ブロック・サーブ・レシーブそれぞれの1本ごとの記録（誰が・どの評価で・
    壁時計の何時何分何秒か）を全部リストにしておく。動画側の秒数への変換は、
    試合ごとのVIDEO_OFFSET_SECONDSを使って表示側（HTML/JS）で行う
    （動画のズレが分かったときに、この関数をやり直さずに直せるように）。
    """
    plays = []
    for line in lines_in_set:
        fields = line.split(';')
        code = fields[0]
        m = ACTION_RE.match(code)
        if not m:
            continue
        team, player_num, skill, skill_type, evaluation, combo = m.groups()
        if team != own_marker:
            continue
        if len(fields) <= 7 or not fields[7]:
            continue
        try:
            wallclock = parse_wallclock_seconds(fields[7])
        except ValueError:
            continue
        plays.append({
            'number': int(player_num), 'skill': skill, 'evaluation': evaluation,
            'wallclock': wallclock,
        })
    return plays


def merge_stats(all_set_stats):
    """複数セット分の集計を、選手ごとに合計する"""
    total = {}
    for set_stats in all_set_stats:
        for number, s in set_stats.items():
            if number not in total:
                total[number] = dict(EMPTY_PLAYER_STATS)
            for key in s:
                total[number][key] += s[key]
    return total


def get_match_date(content):
    """[3MATCH]セクションの1行目の先頭 = 試合の日付"""
    match_section = get_section(content, '3MATCH')
    first_line = match_section.splitlines()[0]
    return first_line.split(';')[0]


def date_slug(match_date):
    """試合日('19/07/2026'のようなDD/MM/YYYY形式)を、並び替えやすい'2026-07-19'の形にする"""
    try:
        day, month, year = match_date.split('/')
        return f'{year}-{month}-{day}'
    except ValueError:
        return match_date.replace('/', '-')


def build_match_slug(match_date, opponent):
    """試合日＋対戦相手から、ファイル名などに使う識別子を作る（拡張子なし）"""
    return f'{date_slug(match_date)}_{opponent}'


def save_match_data(dashboard_data, team, match_date, opponent):
    """
    1試合分のダッシュボードデータを、HTMLではなくJSONファイルとして保存する。
    例: women/2026-07-19_江戸川大学.json
    HTMLは女子・男子それぞれ1枚の共通ダッシュボード(index.html)だけを使い、
    その中で試合を切り替えると、この対応するJSONを読み込みに行く仕組みにしている。
    （GitHubへのアップロードをシンプルにするため、サブフォルダを作らず
    index.htmlと同じ階層にそのまま置く形にしている）
    戻り値は、index.htmlから見た相対パス（例: '2026-07-19_江戸川大学.json'）
    """
    slug = build_match_slug(match_date, opponent)
    relative_path = f'{slug}.json'
    full_path = os.path.join(team, relative_path)
    folder = os.path.dirname(full_path)
    if folder:
        os.makedirs(folder, exist_ok=True)
    with open(full_path, 'w', encoding='utf-8') as f:
        json.dump(dashboard_data, f, ensure_ascii=False)
    return relative_path.replace(os.sep, '/')


def update_manifest(team, match_date, opponent, data_relpath):
    """
    そのチームの試合一覧(matches.json)を更新する。
    同じ試合日・対戦相手の組み合わせが既にあれば古い方を消してから追加するので、
    同じ試合を後で読み込み直しても重複しない。日付が新しい順に並べ直す。
    """
    manifest_path = os.path.join(team, 'matches.json')
    if os.path.exists(manifest_path):
        with open(manifest_path, encoding='utf-8') as f:
            matches = json.load(f)
    else:
        matches = []

    matches = [
        m for m in matches
        if not (m['matchDate'] == match_date and m['opponent'] == opponent)
    ]
    matches.append({
        'matchDate': match_date,
        'opponent': opponent,
        'dataFile': data_relpath,
    })
    matches.sort(key=lambda m: date_slug(m['matchDate']), reverse=True)

    with open(manifest_path, 'w', encoding='utf-8') as f:
        json.dump(matches, f, ensure_ascii=False, indent=2)
    return matches


def rate(numerator, denominator):
    """0割りを避けつつ割合を計算する（分母が0ならNoneを返す）"""
    if denominator == 0:
        return None
    return numerator / denominator


def player_summary(number, s, name=None):
    """1人分の集計データに、名前と決定率・効果率を計算して付け加える"""
    attempts = s['attempts']
    points = s['points']
    errors = s['errors']
    blocked = s['blocked']
    result = {
        'number': number,
        'name': name if name is not None else ROSTER.get(number, f'#{number}'),
        'attempts': attempts,
        'points': points,
        'errors': errors,
        'blocked': blocked,
        'kill_rate': rate(points, attempts),
        'efficiency': rate(points - errors - blocked, attempts),
        'back_attempts': s['back_attempts'],
        'back_points': s['back_points'],
        'back_errors': s['back_errors'],
        'block_attempts': s['block_attempts'],
        'block_points': s['block_points'],
        'block_errors': s['block_errors'],
        'block_touch_own': s['block_touch_own'],
        'block_touch_opp': s['block_touch_opp'],
        'serve_attempts': s['serve_attempts'],
        'serve_aces': s['serve_aces'],
        'serve_errors': s['serve_errors'],
        'serve_half_credit': s['serve_half_credit'],
        'serve_ace_rate': rate(s['serve_aces'], s['serve_attempts']),
        'serve_error_rate': rate(s['serve_errors'], s['serve_attempts']),
        'serve_efficiency': rate(s['serve_aces'] - s['serve_errors'] + 0.5 * s['serve_half_credit'], s['serve_attempts']),
        'receive_attempts': s['receive_attempts'],
        'receive_a': s['receive_a'],
        'receive_b': s['receive_b'],
        'receive_c': s['receive_c'],
        'receive_d': s['receive_d'],
        'receive_errors': s['receive_errors'],
        'receive_return_rate': rate(s['receive_attempts'] - s['receive_errors'], s['receive_attempts']),
        'receive_a_rate': rate(s['receive_a'], s['receive_attempts']),
        'receive_ab_rate': rate(s['receive_a'] + s['receive_b'], s['receive_attempts']),
    }
    for tech in TECH_ORDER:
        t_attempts = s[f'tech_{tech}_attempts']
        t_points = s[f'tech_{tech}_points']
        t_errors = s[f'tech_{tech}_errors']
        result[f'tech_{tech}_attempts'] = t_attempts
        result[f'tech_{tech}_points'] = t_points
        result[f'tech_{tech}_errors'] = t_errors
        result[f'tech_{tech}_kill_rate'] = rate(t_points, t_attempts)
    return result


def team_stats(stats_dict):
    """選手ごとの集計(stats_dict)を、チーム全体の合計に足し合わせる"""
    total = dict(EMPTY_PLAYER_STATS)
    for s in stats_dict.values():
        for key in total:
            total[key] += s[key]
    return total


def build_dashboard_data(all_set_stats, total_stats, match_date, all_rotation_stats, all_side_out_break,
                          all_score_progression, all_side_out_break_by_rotation=None,
                          all_player_rotation_attacks=None, all_starting_lineups=None,
                          all_play_logs=None, all_opp_side_out_break_by_rotation=None,
                          all_opp_rotation_stats=None):
    """ダッシュボード(HTML)に埋め込むためのデータをまとめる"""
    numbers = sorted(total_stats.keys())
    merged_sob_by_rotation = (
        merge_side_out_break_by_rotation(all_side_out_break_by_rotation)
        if all_side_out_break_by_rotation else None
    )
    merged_opp_sob_by_rotation = (
        merge_side_out_break_by_rotation(all_opp_side_out_break_by_rotation)
        if all_opp_side_out_break_by_rotation else None
    )
    player_rotation_summary = (
        build_player_rotation_summary(merge_player_rotation_attacks(all_player_rotation_attacks))
        if all_player_rotation_attacks else {}
    )
    return {
        'matchDate': match_date,
        'opponent': OPPONENT_NAME,
        'teamLabel': TEAM_LABEL,
        'sets': [
            {
                'setNumber': i + 1,
                'players': [player_summary(n, s) for n, s in sorted(set_stats.items())],
            }
            for i, set_stats in enumerate(all_set_stats)
        ],
        'total': [player_summary(n, total_stats[n]) for n in numbers],
        'team': {
            'name': 'チーム全体',
            'total': player_summary(None, team_stats(total_stats), name='チーム全体'),
            'sets': [
                {'setNumber': i + 1, **player_summary(None, team_stats(set_stats), name='チーム全体')}
                for i, set_stats in enumerate(all_set_stats)
            ],
        },
        'rotation': {
            'bySet': [
                {'setNumber': i + 1, 'rows': rotation_summary(rs)}
                for i, rs in enumerate(all_rotation_stats)
            ],
            'total': {'rows': rotation_summary(merge_rotation_stats(all_rotation_stats))},
            'byPlayer': player_rotation_summary,
        },
        'sideOutBreak': {
            'bySet': [
                {'setNumber': i + 1, **side_out_break_summary(r)}
                for i, r in enumerate(all_side_out_break)
            ],
            'total': side_out_break_summary(merge_side_out_break(all_side_out_break)),
            'byRotation': (
                [
                    {'rotation': n, **side_out_break_summary(merged_sob_by_rotation[n])}
                    for n in range(1, 7)
                ]
                if merged_sob_by_rotation else []
            ),
        },
        'opponentSideOutBreak': {
            'byRotation': (
                [
                    {'rotation': n, **side_out_break_summary(merged_opp_sob_by_rotation[n])}
                    for n in range(1, 7)
                ]
                if merged_opp_sob_by_rotation else []
            ),
        },
        'opponentRotation': {
            'bySet': [
                {'setNumber': i + 1, 'rows': rotation_summary(rs)}
                for i, rs in enumerate(all_opp_rotation_stats)
            ] if all_opp_rotation_stats else [],
            'total': {
                'rows': rotation_summary(merge_rotation_stats(all_opp_rotation_stats))
            } if all_opp_rotation_stats else {'rows': []},
        },
        'scoreProgression': {
            'bySet': [
                {
                    'setNumber': i + 1,
                    'points': points,
                    'startingLineup': (
                        all_starting_lineups[i] if all_starting_lineups and i < len(all_starting_lineups) else None
                    ),
                }
                for i, points in enumerate(all_score_progression)
            ],
        },
        'video': {
            'url': VIDEO_URL,
            'offsetSeconds': VIDEO_OFFSET_SECONDS,
            'plays': (
                [
                    {'setNumber': i + 1, **p}
                    for i, set_plays in enumerate(all_play_logs)
                    for p in set_plays
                ]
                if all_play_logs else []
            ),
        },
    }


def generate_dashboard_shell(team, team_label):
    """
    チーム共通のダッシュボード(HTML)を1枚だけ作る。
    このHTMLには試合データを埋め込まず、matches.json と data/*.json を
    ブラウザ側でfetchして表示する仕組みになっている。
    毎回上書きするので、テンプレートを直しても既存の試合データはそのまま使える。
    """
    html = HTML_TEMPLATE.replace('__TEAM_LABEL__', team_label)
    html = html.replace('__PAGE_TITLE__', PAGE_TITLES[team])
    html = html.replace('__OWN_TEAM_CODE__', TARGET_TEAM_CODE_BY_TEAM[team])
    html = html.replace('__ROSTER_JSON__', json.dumps(ROSTER_BY_TEAM[team], ensure_ascii=False))
    output_path = os.path.join(team, 'index.html')
    folder = os.path.dirname(output_path)
    if folder:
        os.makedirs(folder, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    return output_path


def print_report(all_set_stats, total_stats, all_rotation_stats, all_side_out_break):
    print('=' * 70)
    print('スパイクスタッツ自動集計結果')
    print('=' * 70)

    for set_num, set_stats in enumerate(all_set_stats, start=1):
        print(f'\n--- 第{set_num}セット ---')
        print_table(set_stats)

    print('\n--- 全セット合計 ---')
    print_table(total_stats)

    print('\n' + '=' * 70)
    print('ブロックスタッツ自動集計結果（全セット合計）')
    print('=' * 70)
    print_block_table(total_stats)

    print('\n' + '=' * 70)
    print('サーブスタッツ自動集計結果（全セット合計）')
    print('=' * 70)
    print_serve_table(total_stats)

    print('\n' + '=' * 70)
    print('レシーブスタッツ自動集計結果（全セット合計）')
    print('=' * 70)
    print_receive_table(total_stats)

    print('\n' + '=' * 70)
    print('ローテーション別 攻撃タイプ分布自動集計結果（チーム全体）')
    print('=' * 70)
    for set_num, rs in enumerate(all_rotation_stats, start=1):
        print_rotation_table(rotation_summary(rs), f'第{set_num}セット')
    print_rotation_table(rotation_summary(merge_rotation_stats(all_rotation_stats)), '全セット合計')

    print('\n' + '=' * 70)
    print('サイドアウト率・ブレイク率自動集計結果（チーム全体）')
    print('=' * 70)
    header = f"{'':<10}{'サーブ本数':>10}{'ブレイク率':>10}{'レシーブ本数':>12}{'サイドアウト率':>14}"
    print(header)
    print('-' * len(header))
    for set_num, r in enumerate(all_side_out_break, start=1):
        sb = side_out_break_summary(r)
        print(f"{f'第{set_num}セット':<10}{sb['serve_total']:>10}{fmt_pct(sb['break_rate']):>10}"
              f"{sb['receive_total']:>12}{fmt_pct(sb['side_out_rate']):>14}")
    total_sb = side_out_break_summary(merge_side_out_break(all_side_out_break))
    print(f"{'合計':<10}{total_sb['serve_total']:>10}{fmt_pct(total_sb['break_rate']):>10}"
          f"{total_sb['receive_total']:>12}{fmt_pct(total_sb['side_out_rate']):>14}")


def print_table(stats_dict):
    header = f"{'選手':<10}{'打数':>5}{'得点':>5}{'ミス':>5}{'決定率':>8}{'効果率':>8}{'内バック':>9}"
    print(header)
    print('-' * len(header))
    # 背番号順に並べる
    for number in sorted(stats_dict.keys()):
        s = stats_dict[number]
        name = ROSTER.get(number, f'#{number}')
        attempts = s['attempts']
        points = s['points']
        errors = s['errors']
        blocked = s['blocked']

        if attempts > 0:
            kill_rate = points / attempts
            efficiency = (points - errors - blocked) / attempts
            kill_rate_str = f'{kill_rate:.1%}'
            efficiency_str = f'{efficiency:.1%}'
        else:
            kill_rate_str = '-'
            efficiency_str = '-'

        back = f"{s['back_points']}/{s['back_attempts']}"
        print(f"{name:<10}{attempts:>5}{points:>5}{errors:>5}{kill_rate_str:>8}{efficiency_str:>8}{back:>9}")


def print_block_table(stats_dict):
    header = f"{'選手':<10}{'本数':>5}{'得点':>5}{'ワンチ':>7}{'自コート':>7}{'相手コート':>9}{'失点':>5}"
    print(header)
    print('-' * len(header))
    for number in sorted(stats_dict.keys()):
        s = stats_dict[number]
        name = ROSTER.get(number, f'#{number}')
        touch = s['block_touch_own'] + s['block_touch_opp']
        if s['block_attempts'] == 0:
            continue
        print(f"{name:<10}{s['block_attempts']:>5}{s['block_points']:>5}{touch:>7}"
              f"{s['block_touch_own']:>7}{s['block_touch_opp']:>9}{s['block_errors']:>5}")


def print_serve_table(stats_dict):
    header = f"{'選手':<10}{'打数':>5}{'エース':>7}{'ミス':>5}{'エース率':>9}{'ミス率':>8}{'効果率':>8}"
    print(header)
    print('-' * len(header))
    for number in sorted(stats_dict.keys()):
        s = stats_dict[number]
        name = ROSTER.get(number, f'#{number}')
        if s['serve_attempts'] == 0:
            continue
        ace_rate = s['serve_aces'] / s['serve_attempts']
        err_rate = s['serve_errors'] / s['serve_attempts']
        efficiency = (s['serve_aces'] - s['serve_errors'] + 0.5 * s['serve_half_credit']) / s['serve_attempts']
        print(f"{name:<10}{s['serve_attempts']:>5}{s['serve_aces']:>7}{s['serve_errors']:>5}"
              f"{ace_rate:>9.1%}{err_rate:>8.1%}{efficiency:>8.1%}")


def print_receive_table(stats_dict):
    header = (f"{'選手':<10}{'本数':>5}{'A':>5}{'B':>5}{'C':>5}{'D':>5}"
              f"{'ミス':>5}{'返球率':>8}{'A率':>7}{'AB率':>8}")
    print(header)
    print('-' * len(header))
    for number in sorted(stats_dict.keys()):
        s = stats_dict[number]
        name = ROSTER.get(number, f'#{number}')
        if s['receive_attempts'] == 0:
            continue
        return_rate = (s['receive_attempts'] - s['receive_errors']) / s['receive_attempts']
        a_rate = s['receive_a'] / s['receive_attempts']
        ab_rate = (s['receive_a'] + s['receive_b']) / s['receive_attempts']
        print(f"{name:<10}{s['receive_attempts']:>5}{s['receive_a']:>5}{s['receive_b']:>5}"
              f"{s['receive_c']:>5}{s['receive_d']:>5}{s['receive_errors']:>5}{return_rate:>8.1%}{a_rate:>7.1%}{ab_rate:>8.1%}")


def fmt_pct(v):
    return '-' if v is None else f'{v:.1%}'


def print_rotation_table(rows, title):
    """rotation_summary()が返す行リストから、ローテーション×攻撃タイプの表を表示する"""
    print(f'\n--- {title} ---')
    header = f"{'ローテ':<6}{'打数':>5}{'得点':>5}{'ミス':>5}{'決定率':>7}"
    for cat in CATEGORY_ORDER:
        header += f"{CATEGORY_LABELS[cat] + '本数':>9}{'決定率':>7}"
    print(header)
    print('-' * len(header))
    for row in rows:
        label = f"S{row['rotation']}" if isinstance(row['rotation'], int) else row['rotation']
        line = f"{label:<6}{row['attempts']:>5}{row['points']:>5}{row['errors']:>5}{fmt_pct(row['kill_rate']):>7}"
        for cat in CATEGORY_ORDER:
            c = row['categories'][cat]
            line += f"{c['attempts']:>9}{fmt_pct(c['kill_rate']):>7}"
        print(line)


def main():
    content = load_file(FILE_PATH)
    own_marker = get_own_team_marker(content)
    back_attack_combos = get_back_attack_combos(content)
    combo_categories = get_combo_categories(content)
    match_date = get_match_date(content)

    scout_section = get_section(content, '3SCOUT')
    # ※以前はここで l.split(';')[0] として行の後半(得点・ローテーション情報)を
    # 切り捨てていましたが、ローテーション別攻撃分布の集計に必要なので
    # 行全体をそのまま残すように変更しました
    all_lines = [l for l in scout_section.splitlines() if l.strip()]

    sets_lines = split_into_sets(all_lines)
    all_set_stats = [analyze_set(s, own_marker, back_attack_combos) for s in sets_lines]
    total_stats = merge_stats(all_set_stats)
    all_rotation_stats = [analyze_rotation_attacks(s, own_marker, combo_categories) for s in sets_lines]
    all_side_out_break = [analyze_side_out_break(s, own_marker) for s in sets_lines]
    all_side_out_break_by_rotation = [analyze_side_out_break_by_rotation(s, own_marker) for s in sets_lines]
    all_player_rotation_attacks = [analyze_player_rotation_attacks(s, own_marker) for s in sets_lines]
    all_score_progression = [analyze_score_progression(s, own_marker) for s in sets_lines]
    all_starting_lineups = [get_starting_lineup(s, own_marker) for s in sets_lines]
    all_play_logs = [build_play_log(s, own_marker) for s in sets_lines]

    # 相手チーム分析（スカウティング）用：ownとoppを入れ替えて同じ関数を再利用するだけで、
    # 相手チーム自身のローテーション番号でのサイドアウト率・ブレイク率が計算できる
    # （.dvwには両チームのローテーション情報が常に両方入っているため）
    opponent_marker = 'a' if own_marker == '*' else '*'
    all_opp_side_out_break_by_rotation = [
        analyze_side_out_break_by_rotation(s, opponent_marker) for s in sets_lines
    ]
    # 相手が決めた攻撃のコース（レフト/ミドル/ライト）用：同じくownとoppを入れ替えるだけで、
    # 相手自身のコンビ分類の攻撃タイプ内訳が計算できる（2026-09-03追加）
    all_opp_rotation_stats = [
        analyze_rotation_attacks(s, opponent_marker, combo_categories) for s in sets_lines
    ]

    print_report(all_set_stats, total_stats, all_rotation_stats, all_side_out_break)

    dashboard_data = build_dashboard_data(
        all_set_stats, total_stats, match_date, all_rotation_stats, all_side_out_break,
        all_score_progression, all_side_out_break_by_rotation, all_player_rotation_attacks,
        all_starting_lineups, all_play_logs, all_opp_side_out_break_by_rotation,
        all_opp_rotation_stats)

    data_relpath = save_match_data(dashboard_data, TEAM, match_date, OPPONENT_NAME)
    matches = update_manifest(TEAM, match_date, OPPONENT_NAME, data_relpath)
    shell_path = generate_dashboard_shell(TEAM, TEAM_LABEL)

    print(f'\n試合データを保存しました: {os.path.join(TEAM, data_relpath)}')
    print(f'ダッシュボード（共通1枚）を更新しました: {shell_path}')
    print(f'このチームの試合一覧: {len(matches)}試合')
    print(f'\n{TEAM}フォルダの中身（index.html, matches.json, data/フォルダ全部）を'
          f'GitHubへアップロードしてください。以前アップロードした古いindex.htmlは上書きされます。')


if __name__ == '__main__':
    main()
