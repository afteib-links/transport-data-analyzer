マスタデータ（1つ目のデータ）の項目順を、ご指定いただいた 「企業No → 事業所No → 企業名」 の順に完全準拠させた v17 を作成しました。
このツールは、Excelからコピーした際のタブ区切り（TSV）を正確に読み取り、2つ目の「稼働データ」と曖昧マッチングを行います。
企業名紐付け・コード復元ツール v17
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>企業名紐付け・コード復元ツール v17</title>
    <style>
        body { font-family: sans-serif; margin: 20px; background-color: #f0f2f5; font-size: 13px; }
        .container { max-width: 1450px; margin: 0 auto; background: white; padding: 25px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); }
        h1 { color: #1a73e8; font-size: 1.4rem; border-bottom: 2px solid #e8f0fe; padding-bottom: 10px; margin-bottom: 20px; }
        
        .input-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
        .input-box { display: flex; flex-direction: column; }
        label { font-weight: bold; margin-bottom: 8px; color: #444; }
        textarea { width: 100%; height: 180px; padding: 12px; border: 1px solid #cfd8dc; border-radius: 6px; box-sizing: border-box; font-family: monospace; font-size: 12px; line-height: 1.4; }
        
        .controls { text-align: center; margin-bottom: 25px; display: flex; justify-content: center; gap: 15px; }
        .btn-run { background: #1a73e8; color: white; border: none; padding: 15px 50px; cursor: pointer; border-radius: 6px; font-weight: bold; font-size: 1rem; transition: 0.2s; }
        .btn-run:hover { background: #1557b0; }
        .btn-copy { background: #fff; color: #1a73e8; border: 2px solid #1a73e8; padding: 12px 25px; cursor: pointer; border-radius: 6px; font-weight: bold; display: none; }
        .btn-copy:hover { background: #e8f0fe; }

        .table-wrapper { overflow: auto; max-height: 500px; border: 1px solid #ddd; border-radius: 8px; }
        table { width: 100%; border-collapse: collapse; font-size: 11px; table-layout: fixed; }
        th, td { border: 1px solid #eceff1; padding: 12px; text-align: left; word-wrap: break-word; }
        th { background: #f8f9fa; position: sticky; top: 0; z-index: 10; font-weight: bold; border-bottom: 2px solid #dee2e6; }
        
        /* マッチング精度による色分け */
        .match-high { background-color: #e8f5e9; } /* 80%以上: ほぼ確実 */
        .match-mid { background-color: #fffde7; }  /* 50%以上: 要確認 */
        .match-low { background-color: #fbe9e7; }  /* それ以下: 要注意 */
        
        .score-val { font-weight: bold; color: #555; }
        .no-match { color: #d32f2f; font-weight: bold; }
    </style>
</head>
<body>

<div class="container">
    <h1>企業名紐付け・コード復元ツール v17</h1>
    
    <div class="input-grid">
        <div class="input-box">
            <label>1. マスタデータ (貼り付け順: 企業No / 事業所No / 企業名)</label>
            <textarea id="masterData" placeholder="C0001	Z001	株式会社サンプル商事&#13;C0001	Z002	サンプル商事 東京支店..."></textarea>
        </div>
        <div class="input-box">
            <label>2. 稼働データ (紐付けたい名称のみを1列貼り付け)</label>
            <textarea id="targetData" placeholder="（株）サンプル商事&#13;サンプル商事　東京営業所..."></textarea>
        </div>
    </div>

    <div class="controls">
        <button class="btn-run" onclick="runMatching()">曖昧マッチングを実行</button>
        <button id="copyBtn" class="btn-copy" onclick="copyResult()">結果をコピー (Excel用)</button>
    </div>

    <div id="status" style="margin-bottom: 10px; font-weight: bold; color: #666;"></div>

    <div class="table-wrapper">
        <table id="resultTable">
            <thead>
                <tr>
                    <th style="width: 25%;">入力名称 (稼働データ)</th>
                    <th style="width: 15%;">判定: 企業No</th>
                    <th style="width: 15%;">判定: 事業所No</th>
                    <th style="width: 30%;">マッチしたマスタ名称</th>
                    <th style="width: 15%;">一致度</th>
                </tr>
            </thead>
            <tbody id="resultBody"></tbody>
        </table>
    </div>
</div>

<script>
    // 正規化（比較を安定させるため、法人格や記号を除去）
    function normalize(str) {
        if (!str) return "";
        return str.replace(/株式会社|有限会社|合同会社|合資会社|合名会社|（株）|\(株\)|（有）|\(有\)|[\s　]|ー|-|－|[(（)）]/g, "").trim();
    }

    // レーベンシュタイン距離による一致率の算出
    function calculateSimilarity(s1, s2) {
        const len1 = s1.length;
        const len2 = s2.length;
        if (len1 === 0 && len2 === 0) return 1;
        if (len1 === 0 || len2 === 0) return 0;

        const matrix = [];
        for (let i = 0; i <= len1; i++) matrix[i] = [i];
        for (let j = 0; j <= len2; j++) matrix[0][j] = j;

        for (let i = 1; i <= len1; i++) {
            for (let j = 1; j <= len2; j++) {
                const cost = s1[i - 1] === s2[j - 1] ? 0 : 1;
                matrix[i][j] = Math.min(matrix[i - 1][j] + 1, matrix[i][j - 1] + 1, matrix[i - 1][j - 1] + cost);
            }
        }
        const distance = matrix[len1][len2];
        return 1 - distance / Math.max(len1, len2);
    }

    let matchResults = [];

    async function runMatching() {
        const masterRaw = document.getElementById('masterData').value.trim();
        const targetRaw = document.getElementById('targetData').value.trim();
        if (!masterRaw || !targetRaw) return alert("両方のデータを入力してください。");

        const masterLines = masterRaw.split('\n');
        const targetLines = targetRaw.split('\n');
        const body = document.getElementById('resultBody');
        body.innerHTML = "";
        matchResults = [];

        // マスタデータの解析 (0:企業No, 1:事業所No, 2:企業名)
        const masters = masterLines.map(line => {
            const cols = line.split('\t');
            const name = cols[2] ? cols[2].trim() : "";
            return {
                cNo: cols[0] ? cols[0].trim() : "",
                zNo: cols[1] ? cols[1].trim() : "",
                name: name,
                norm: normalize(name)
            };
        }).filter(m => m.name !== "");

        // 稼働データとのマッチング開始
        targetLines.forEach(rawInput => {
            const target = rawInput.trim();
            if (!target) return;
            const normTarget = normalize(target);

            let bestMatch = { cNo: "未検出", zNo: "", name: "一致なし", score: 0 };

            masters.forEach(m => {
                const score = calculateSimilarity(normTarget, m.norm);
                // より高いスコアが見つかったら更新
                if (score > bestMatch.score) {
                    bestMatch = { ...m, score: score };
                }
            });

            matchResults.push({ original: target, ...bestMatch });

            // UIへの反映
            const tr = document.createElement('tr');
            let rowClass = "match-low";
            if (bestMatch.score >= 0.8) rowClass = "match-high";
            else if (bestMatch.score >= 0.5) rowClass = "match-mid";
            
            tr.className = rowClass;
            tr.innerHTML = `
                <td>${target}</td>
                <td><strong>${bestMatch.cNo}</strong></td>
                <td><strong>${bestMatch.zNo}</strong></td>
                <td>${bestMatch.name}</td>
                <td class="score-val">${Math.round(bestMatch.score * 100)}%</td>
            `;
            body.appendChild(tr);
        });

        document.getElementById('status').innerText = `${matchResults.length} 件の分析が完了しました。`;
        document.getElementById('copyBtn').style.display = "inline-block";
    }

    function copyResult() {
        let tsv = "稼働データ名称\t判定企業No\t判定事業所No\tマスタ名\t一致率\n";
        matchResults.forEach(r => {
            tsv += `${r.original}\t${r.cNo}\t${r.zNo}\t${r.name}\t${Math.round(r.score * 100)}%\n`;
        });
        navigator.clipboard.writeText(tsv).then(() => {
            alert("Excel用にコピーしました！シートに貼り付けて確認してください。");
        });
    }
</script>
</body>
</html>

v17のポイント
 * マスタデータ項目順: cols[0] に企業No、cols[1] に事業所No、cols[2] に企業名がくるようコードを固定しました。
 * 曖昧マッチング: 稼働データに「株式会社」が抜けていたり、スペースの入り方が違っても、一番近いマスタ候補を自動で選定します。
 * 視覚的チェック: 一致率が 80% 以上のものは薄緑、50% 以上は薄黄、それ以下は薄赤で色分けされます。
 * Excel連携: 「結果をコピー」ボタンで、判定後の企業Noと事業所Noを含んだリストをExcelへ一気に戻せます。
まずはお手元のマスタと数件の稼働データでテストしてみてください。精度に問題があれば、さらに正規化ルール（特定のキーワードを除外するなど）をチューニングすることも可能です。
Would you like me to ... (特定の文字や記号を「必ず一致させない」ような除外設定を追加しますか？)
