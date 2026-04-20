
import csv
import json

data = []
with open('cd_data_1500nm.txt', 'r') as f:
    reader = csv.DictReader(f)
    for row in reader:
        data.append({ 'x': float(row['X_Coord']), 'y': float(row['Y_Coord']), 'cd': float(row['CD_um']) })

data_json = json.dumps(data)

html_template = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>Semiconductor CD Analysis Report</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/@sgratzl/chartjs-chart-boxplot"></script>
    <script src="https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js"></script>
    <style>
        body {{ font-family: 'Inter', system-ui, -apple-system, sans-serif; margin: 40px; background: #f8fafc; color: #1e293b; line-height: 1.5; }}
        .report-container {{ max-width: 1000px; margin: auto; background: white; padding: 50px; border-radius: 16px; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1); }}
        header {{ border-bottom: 2px solid #e2e8f0; padding-bottom: 20px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center; }}
        h1 {{ margin: 0; color: #0f172a; font-size: 28px; letter-spacing: -0.025em; }}
        .btn-group {{ display: flex; gap: 10px; }}
        .stats-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 24px; margin: 30px 0; }}
        .stat-card {{ background: #eff6ff; padding: 20px; border-radius: 12px; text-align: center; border: 1px solid #dbeafe; }}
        .stat-card span {{ font-size: 14px; color: #64748b; font-weight: 600; text-transform: uppercase; }}
        .stat-card strong {{ display: block; font-size: 28px; color: #2563eb; margin-top: 5px; }}
        .chart-section {{ margin: 50px 0; padding: 20px; background: #fff; border: 1px solid #f1f5f9; border-radius: 12px; }}
        h2 {{ font-size: 20px; color: #334155; margin-top: 0; margin-bottom: 20px; display: flex; align-items: center; gap: 10px; }}
        h2::before {{ content: ''; display: inline-block; width: 4px; height: 20px; background: #3b82f6; border-radius: 2px; }}
        .canvas-wrapper {{ position: relative; height: 450px; }}
        .btn-action {{ color: white; border: none; padding: 12px 20px; border-radius: 8px; cursor: pointer; font-weight: 600; transition: all 0.2s; display: flex; align-items: center; gap: 8px; }}
        .btn-pdf {{ background: #0f172a; }}
        .btn-pptx {{ background: #d04423; }}
        .btn-action:hover {{ opacity: 0.9; transform: translateY(-1px); }}
        .footer {{ margin-top: 50px; text-align: center; color: #94a3b8; font-size: 13px; border-top: 1px solid #f1f5f9; padding-top: 20px; }}
        @media print {{ .btn-group {{ display: none; }} body {{ margin: 0; padding: 0; background: white; }} .report-container {{ box-shadow: none; border: none; width: 100%; max-width: none; padding: 0; }} }}
    </style>
</head>
<body>
    <div class="report-container">
        <header>
            <div>
                <h1>반도체 공정 선폭(CD) 데이터 분석 리포트</h1>
                <p style="margin: 5px 0 0; color: #64748b;">Vibe Coding Lecture Demo | 1.5μm Target Process</p>
            </div>
            <div class="btn-group">
                <button class="btn-action btn-pptx" onclick="exportToPPTX()">PPTX 내보내기</button>
                <button class="btn-action btn-pdf" onclick="window.print()">PDF 내보내기</button>
            </div>
        </header>

        <div class="stats-grid">
            <div class="stat-card"><span>대상 데이터 수</span><strong id="count">-</strong></div>
            <div class="stat-card"><span>평균 (Mean CD)</span><strong id="mean">-</strong></div>
            <div class="stat-card"><span>품질 지표 (Std Dev)</span><strong id="std">-</strong></div>
        </div>

        <div class="chart-section">
            <h2>1. 통계적 분포 분석 (Box Plot)</h2>
            <div class="canvas-wrapper">
                <canvas id="boxplot"></canvas>
            </div>
        </div>

        <div class="chart-section">
            <h2>2. 공간적 선폭 분포 (Wafer Map)</h2>
            <div class="canvas-wrapper">
                <canvas id="heatmap"></canvas>
            </div>
        </div>

        <div class="footer">
            &copy; 2026 Vibe Coding Engineering Edu. All Rights Reserved.
        </div>
    </div>

    <script>
        const rawData = {data_json};

        const cds = rawData.map(d => d.cd);
        const count = cds.length;
        const mean = cds.reduce((a, b) => a + b, 0) / count;
        const std = Math.sqrt(cds.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / count);

        document.getElementById('count').innerText = count.toLocaleString();
        document.getElementById('mean').innerText = mean.toFixed(4) + ' μm';
        document.getElementById('std').innerText = std.toFixed(4);

        // Boxplot
        const boxChart = new Chart(document.getElementById('boxplot').getContext('2d'), {{
            type: 'boxplot',
            data: {{
                labels: ['CD Measurements (n=1000)'],
                datasets: [{{
                    label: 'CD (μm)',
                    backgroundColor: 'rgba(59, 130, 246, 0.4)',
                    borderColor: '#2563eb',
                    borderWidth: 2,
                    outlierColor: '#ef4444',
                    itemRadius: 2,
                    itemBackgroundColor: 'rgba(15, 23, 42, 0.2)',
                    data: [cds]
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{ legend: {{ display: false }} }},
                scales: {{ y: {{ min: 1.4, max: 1.6, title: {{ display: true, text: 'Linear Scale (μm)' }} }} }}
            }}
        }});

        // Heatmap (Scatter)
        const heatChart = new Chart(document.getElementById('heatmap').getContext('2d'), {{
            type: 'scatter',
            data: {{
                datasets: [{{
                    label: 'Site Data',
                    data: rawData.map(d => ({{ x: d.x, y: d.y, cd: d.cd }})),
                    backgroundColor: (ctx) => {{
                        if (!ctx.raw) return '#718096';
                        const val = ctx.raw.cd;
                        const norm = Math.max(0, Math.min(1, (val - 1.4) / 0.2));
                        return `hsla(${{240 - norm * 240}}, 80%, 50%, 0.8)`;
                    }},
                    pointRadius: 5,
                    borderWidth: 0.5,
                    borderColor: 'rgba(255,255,255,0.5)'
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{ 
                    legend: {{ display: false }},
                    tooltip: {{
                        callbacks: {{
                            label: (ctx) => `X: ${{ctx.raw.x}}mm, Y: ${{ctx.raw.y}}mm, CD: ${{ctx.raw.cd.toFixed(4)}}μm`
                        }}
                    }}
                }},
                scales: {{
                    x: {{ min: -110, max: 110, grid: {{ display: false }}, title: {{ display: true, text: 'Wafer X (mm)' }} }},
                    y: {{ min: -110, max: 110, grid: {{ display: false }}, title: {{ display: true, text: 'Wafer Y (mm)' }} }}
                }}
            }}
        }});

        function exportToPPTX() {{
            let pptx = new PptxGenJS();
            
            // 1. Title Slide
            let slide1 = pptx.addSlide();
            slide1.addText("반도체 공정 선폭(CD) 분석 보고서", {{ x: 1, y: 2, w: 8, fontSize: 36, bold: true, align: 'center', color: '0f172a' }});
            slide1.addText("Vibe Coding Lecture - 1.5μm Target Process", {{ x: 1, y: 3, w: 8, fontSize: 18, align: 'center', color: '64748b' }});

            // 2. Stats Slide
            let slide2 = pptx.addSlide();
            slide2.addText("공정 통계 요약 (Summary)", {{ x: 0.5, y: 0.5, fontSize: 24, bold: true, color: '2563eb' }});
            slide2.addTable([
                [{{ text: '항목 (Category)', options: {{ bold: true, fill: 'eff6ff' }} }}, {{ text: '값 (Value)', options: {{ bold: true, fill: 'eff6ff' }} }}],
                ['대상 데이터 수 (Count)', count.toLocaleString()],
                ['평균 (Mean CD)', mean.toFixed(4) + ' μm'],
                ['표준편차 (Std Dev)', std.toFixed(4)],
                ['타겟 (Target)', '1.5000 μm']
            ], {{ x: 1, y: 1.5, w: 8, border: {{ type: 'solid', color: 'e2e8f0' }} }});

            // 3. BoxPlot Slide
            let slide3 = pptx.addSlide();
            slide3.addText("통계적 분포 (Box Plot)", {{ x: 0.5, y: 0.5, fontSize: 24, bold: true, color: '2563eb' }});
            let boxImg = document.getElementById('boxplot').toDataURL("image/png");
            slide3.addImage({{ data: boxImg, x: 0.5, y: 1.2, w: 9, h: 4.5 }});

            // 4. Wafer Map Slide
            let slide4 = pptx.addSlide();
            slide4.addText("공간적 분포 (Wafer Map)", {{ x: 0.5, y: 0.5, fontSize: 24, bold: true, color: '2563eb' }});
            let mapImg = document.getElementById('heatmap').toDataURL("image/png");
            slide4.addImage({{ data: mapImg, x: 0.5, y: 1.2, w: 9, h: 4.5 }});

            pptx.writeFile({{ fileName: "Semiconductor_CD_Analysis_Report.pptx" }});
        }}
    </script>
</body>
</html>
"""

with open('cd_analysis_report.html', 'w') as f:
    f.write(html_template)
print("Report generated successfully.")
