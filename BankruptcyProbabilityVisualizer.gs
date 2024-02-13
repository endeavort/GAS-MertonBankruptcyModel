// 破産確率表をスプレッドシートに書き出す関数
function writeRuinRatesToSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // ============================ パラーメーター ============================
  var funds = 1000000; // 初期資金
  var riskRate = 0.01; // リスク率
  var ruinLine = 600000; // 撤退ライン
  var winRange = range(0.1, 1.01, 0.1); // 勝率の範囲
  var rrRange = range(0.2, 3.01, 0.2); // リスクリワード比の範囲
  // ======================================================================

  var data = []; // 結果を格納する配列
  
  rrRange.forEach(function(rr) {
    var row = [];
    winRange.forEach(function(win) {
      var ruinRate = calculateRuinRate(win, rr, riskRate, funds, ruinLine);
      var formattedRate; // フォーマットされた破産確率
      if (ruinRate >= 0.995) {
        formattedRate = "100%";
      } else if (ruinRate < 0.005) {
        formattedRate = "0%";
      } else {
        formattedRate = (ruinRate * 100).toFixed(2) + "%";
      }
      row.push(formattedRate);
    });
    data.push(row);
  });
  
  // A1セルから表を作成
  var startRow = 1; // 開始行
  var startCol = 1; // 開始列
  // 勝率のヘッダーを整数の％表記に変更
  var headerRow = ["Win Rate / RR"].concat(winRange.map(function(win) { return (win * 100).toFixed(0) + "%"; }));
  sheet.getRange(startRow, startCol, 1, headerRow.length).setValues([headerRow]);
  // リスクリワード比の列を小数点第一位に
  var rrCol = rrRange.map(function(rr) { return [rr.toFixed(1)]; });
  sheet.getRange(startRow + 1, startCol, rrCol.length, 1).setValues(rrCol);
  // 結果をスプレッドシートに書き込む
  for (var i = 0; i < data.length; i++) {
    var rowRange = sheet.getRange(i + startRow + 1, startCol + 1, 1, data[i].length);
    rowRange.setValues([data[i]]);
    // セルの色を設定
    data[i].forEach(function(rate, j) {
      var cell = rowRange.getCell(1, j + 1);
      var color = getColorForRate(rate);
      cell.setFontColor(color);
    });
  }
}

// セルの色
function getColorForRate(rate) {
  var value = parseFloat(rate);
  if (rate === "100%") return "#ff0000"; // 赤
  else if (value < 100 && value >= 10) return "#ffa500"; // オレンジ
  else if (value < 10 && value >= 1) return "#ffd700"; // 黄色
  else if (value < 1 && value > 0) return "#adff2f"; // 黄緑
  else if (value === 0) return "#008000"; // 緑
  else return "#000000"; // 黒（該当しない場合）
}

// 破産確率を計算する関数
function calculateRuinRate(winPct, riskReward, riskRate, funds, ruinLine) {
  var a = Math.log(1 + riskReward * riskRate);
  var b = Math.abs(Math.log(1 - riskRate));
  var n = Math.log(funds / ruinLine);
  var R = a / b;
  var S = solveEquation(winPct, R);
  return Math.pow(S, n / b);
}

// 方程式を解く関数
function solveEquation(P, R) {
  var S = 0;
  while (equation(S, P, R) > 0) {
    S += 0.0001;
    if (S >= 1) return 1;
  }
  return S;
}

// 方程式の計算
function equation(x, P, R) {
  return P * Math.pow(x, R + 1) + (1 - P) - x;
}

// 数値範囲生成器
function range(start, stop, step) {
  let a = [];
  for (let b = start; b <= stop; b += step) {
    a.push(b);
  }
  return a;
}

