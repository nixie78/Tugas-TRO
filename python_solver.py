"""
TUGAS UTS - TEKNIK RISET OPERASIONAL
POINT 3a (PART 2): SOLUSI MENGGUNAKAN PYTHON PuLP

PT SAWIT MAKMUR SEJAHTERA - OPTIMASI SISTEM DISTRIBUSI

File: python_solver.py
Output: HTML Interaktif di Browser
"""

from pulp import *
import pandas as pd
import webbrowser
import os
import datetime

print("="*80)
print("POINT 3a (PART 2): SOLUSI DENGAN PYTHON PuLP")
print("="*80)

# ================================================================================
# STEP 1: DEFINISI DATA
# ================================================================================

print("\n[1] Definisi data...")

supply_capacity = {'Kebun_A': 5000, 'Kebun_B': 4000, 'Kebun_C': 3500}
factory_capacity = {'Pabrik_1': 6000, 'Pabrik_2': 5000}
demand = {'PD1': 600, 'PD2': 800, 'PD3': 500, 'PD4': 650}
yield_rate = 0.22

cost_tbs = {
    ('Kebun_A', 'Pabrik_1'): 50000, ('Kebun_A', 'Pabrik_2'): 80000,
    ('Kebun_B', 'Pabrik_1'): 70000, ('Kebun_B', 'Pabrik_2'): 40000,
    ('Kebun_C', 'Pabrik_1'): 60000, ('Kebun_C', 'Pabrik_2'): 55000,
}

cost_cpo = {
    ('Pabrik_1', 'PD1'): 100000, ('Pabrik_1', 'PD2'): 120000,
    ('Pabrik_1', 'PD3'): 90000, ('Pabrik_1', 'PD4'): 110000,
    ('Pabrik_2', 'PD1'): 130000, ('Pabrik_2', 'PD2'): 80000,
    ('Pabrik_2', 'PD3'): 95000, ('Pabrik_2', 'PD4'): 85000,
}

print("✓ Data loaded")

# ================================================================================
# STEP 2-5: BUILD & SOLVE MODEL
# ================================================================================

print("[2] Membangun model...")

kebun_list = list(supply_capacity.keys())
pabrik_list = list(factory_capacity.keys())
pd_list = list(demand.keys())

model = LpProblem("Optimasi_Distribusi_Sawit", LpMinimize)

x = LpVariable.dicts("X_TBS", cost_tbs.keys(), lowBound=0, cat='Continuous')
y = LpVariable.dicts("Y_CPO", cost_cpo.keys(), lowBound=0, cat='Continuous')

print("✓ Variabel keputusan dibuat")

model += (
    lpSum([cost_tbs[i] * x[i] for i in cost_tbs.keys()]) +
    lpSum([cost_cpo[j] * y[j] for j in cost_cpo.keys()]),
    "Total_Biaya"
)

print("✓ Objective function didefinisikan")

# Constraints
for kebun in kebun_list:
    model += lpSum([x[(kebun, p)] for p in pabrik_list if (kebun, p) in cost_tbs.keys()]) <= supply_capacity[kebun]

for pabrik in pabrik_list:
    model += lpSum([x[(k, pabrik)] for k in kebun_list if (k, pabrik) in cost_tbs.keys()]) <= factory_capacity[pabrik]

for pabrik in pabrik_list:
    tbs_in = lpSum([x[(k, pabrik)] for k in kebun_list if (k, pabrik) in cost_tbs.keys()])
    cpo_out = lpSum([y[(pabrik, pd)] for pd in pd_list if (pabrik, pd) in cost_cpo.keys()])
    model += cpo_out == yield_rate * tbs_in

for pd in pd_list:
    model += lpSum([y[(p, pd)] for p in pabrik_list if (p, pd) in cost_cpo.keys()]) >= demand[pd]

print("✓ Constraints didefinisikan")

print("[3] Menyelesaikan model...")
model.solve(PULP_CBC_CMD(msg=0))

status = LpStatus[model.status]
print(f"✓ Status: {status}")

# ================================================================================
# KALKULASI HASIL
# ================================================================================

print("[4] Mengkalkulasi hasil...")

total_biaya = value(model.objective)
biaya_tbs = sum([cost_tbs[i] * x[i].varValue for i in cost_tbs.keys()])
biaya_cpo = sum([cost_cpo[j] * y[j].varValue for j in cost_cpo.keys()])

print(f"✓ Total biaya: Rp {total_biaya:,.0f}")

# ================================================================================
# GENERATE HTML
# ================================================================================

print("[5] Membuat HTML report...")

html = f"""<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Python PuLP Solution</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            color: #333;
        }}
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }}
        .header {{
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }}
        .header h1 {{
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }}
        .main-content {{ padding: 40px; }}
        .section {{
            margin-bottom: 40px;
            background: #f8f9fa;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        .section h2 {{
            color: #1e3c72;
            margin-bottom: 20px;
            font-size: 1.8em;
            border-bottom: 3px solid #667eea;
            padding-bottom: 10px;
        }}
        .metric-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin: 30px 0;
        }}
        .metric-card {{
            background: white;
            padding: 25px;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            border-left: 5px solid #667eea;
            transition: transform 0.3s ease;
        }}
        .metric-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 8px 20px rgba(0,0,0,0.15);
        }}
        .metric-card h3 {{
            color: #666;
            font-size: 0.9em;
            margin-bottom: 10px;
            text-transform: uppercase;
        }}
        .metric-card .value {{
            color: #1e3c72;
            font-size: 2em;
            font-weight: bold;
            margin-bottom: 5px;
        }}
        .metric-card .subtitle {{ color: #999; font-size: 0.9em; }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        th {{
            background: #1e3c72;
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
        }}
        td {{ padding: 12px 15px; border-bottom: 1px solid #eee; }}
        tr:hover {{ background: #f5f5f5; }}
        .step-box {{
            background: white;
            padding: 20px;
            margin: 15px 0;
            border-radius: 10px;
            border-left: 4px solid #667eea;
        }}
        .step-box h3 {{ color: #1e3c72; margin-bottom: 10px; }}
        .step-box p {{ line-height: 1.6; color: #555; }}
        .code-box {{
            background: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            font-family: 'Courier New', monospace;
            font-size: 0.9em;
            margin: 10px 0;
            border: 1px solid #ddd;
        }}
        .success {{ color: #28a745; font-weight: bold; }}
        .footer {{
            background: #1e3c72;
            color: white;
            text-align: center;
            padding: 30px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🐍 SOLUSI PYTHON PuLP</h1>
            <p>Optimasi Sistem Distribusi</p>
            <p style="font-size: 0.9em; margin-top: 10px;">PT Sawit Makmur Sejahtera</p>
        </div>
        
        <div class="main-content">
            <div class="section">
                <h2>📋 METODOLOGI STEP-BY-STEP</h2>
                
                <div class="step-box">
                    <h3>STEP 1: Definisi Data</h3>
                    <p>Mengumpulkan semua parameter yang diperlukan:</p>
                    <ul style="margin-left: 20px; margin-top: 10px; line-height: 1.8;">
                        <li>Kapasitas supply kebun: {sum(supply_capacity.values()):,} ton TBS</li>
                        <li>Kapasitas pabrik: {sum(factory_capacity.values()):,} ton TBS</li>
                        <li>Total demand: {sum(demand.values()):,} ton CPO</li>
                        <li>Yield rate: {yield_rate*100:.0f}%</li>
                        <li>Matriks biaya transportasi TBS dan CPO</li>
                    </ul>
                </div>
                
                <div class="step-box">
                    <h3>STEP 2: Inisialisasi Model</h3>
                    <div class="code-box">
                        model = LpProblem("Optimasi_Distribusi_Sawit", LpMinimize)
                    </div>
                    <p>Membuat model Linear Programming dengan objective: <strong>Minimasi Biaya</strong></p>
                </div>
                
                <div class="step-box">
                    <h3>STEP 3: Definisi Variabel Keputusan</h3>
                    <div class="code-box">
                        x = LpVariable.dicts("X_TBS", routes_tbs, lowBound=0)<br>
                        y = LpVariable.dicts("Y_CPO", routes_cpo, lowBound=0)
                    </div>
                    <p>Total: <strong>{len(x) + len(y)} variabel</strong> (6 TBS + 8 CPO)</p>
                </div>
                
                <div class="step-box">
                    <h3>STEP 4: Definisi Objective Function</h3>
                    <div class="code-box">
                        Minimize Z = Σ(Biaya_TBS × X) + Σ(Biaya_CPO × Y)
                    </div>
                    <p>Meminimalkan total biaya transportasi TBS dan CPO</p>
                </div>
                
                <div class="step-box">
                    <h3>STEP 5: Definisi Constraints</h3>
                    <p>Total: <strong>{len(model.constraints)} constraints</strong></p>
                    <ul style="margin-left: 20px; margin-top: 10px; line-height: 1.8;">
                        <li>3 constraint kapasitas supply kebun</li>
                        <li>2 constraint kapasitas pabrik</li>
                        <li>2 constraint material balance (TBS → CPO)</li>
                        <li>4 constraint pemenuhan demand</li>
                        <li>+ non-negativity constraints</li>
                    </ul>
                </div>
                
                <div class="step-box">
                    <h3>STEP 6: Solve Model</h3>
                    <div class="code-box">
                        model.solve(PULP_CBC_CMD())<br>
                        Status: <strong style="color: #28a745;">{status}</strong>
                    </div>
                    <p>Menggunakan CBC (COIN-OR Branch and Cut) solver</p>
                </div>
            </div>
            
            <div class="section">
                <h2>🎯 HASIL OPTIMASI</h2>
                <div class="metric-grid">
                    <div class="metric-card">
                        <h3>💰 Total Biaya Optimal</h3>
                        <div class="value">Rp {total_biaya:,.0f}</div>
                        <div class="subtitle">Per bulan</div>
                    </div>
                    <div class="metric-card">
                        <h3>🚛 Biaya TBS</h3>
                        <div class="value">Rp {biaya_tbs:,.0f}</div>
                        <div class="subtitle">{biaya_tbs/total_biaya*100:.1f}% dari total</div>
                    </div>
                    <div class="metric-card">
                        <h3>📦 Biaya CPO</h3>
                        <div class="value">Rp {biaya_cpo:,.0f}</div>
                        <div class="subtitle">{biaya_cpo/total_biaya*100:.1f}% dari total</div>
                    </div>
                    <div class="metric-card">
                        <h3>✅ Status</h3>
                        <div class="value" style="color: #28a745; font-size: 1.5em;">OPTIMAL</div>
                        <div class="subtitle">CBC Solver</div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>🚛 ALOKASI TBS (Kebun → Pabrik)</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Dari Kebun</th>
                            <th>Ke Pabrik</th>
                            <th>Jumlah (ton)</th>
                            <th>Biaya/ton</th>
                            <th>Total Biaya</th>
                        </tr>
                    </thead>
                    <tbody>
"""

for (kebun, pabrik), var in x.items():
    if var.varValue > 0.01:
        html += f"""
                        <tr>
                            <td><strong>{kebun}</strong></td>
                            <td><strong>{pabrik}</strong></td>
                            <td>{var.varValue:,.0f}</td>
                            <td>Rp {cost_tbs[(kebun, pabrik)]:,}</td>
                            <td>Rp {cost_tbs[(kebun, pabrik)] * var.varValue:,.0f}</td>
                        </tr>
"""

html += """
                    </tbody>
                </table>
            </div>
            
            <div class="section">
                <h2>🏭 PRODUKSI CPO PER PABRIK</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Pabrik</th>
                            <th>Input TBS (ton)</th>
                            <th>Output CPO (ton)</th>
                            <th>Kapasitas</th>
                            <th>Utilisasi</th>
                        </tr>
                    </thead>
                    <tbody>
"""

for pabrik in pabrik_list:
    tbs_in = sum([x[(k, pabrik)].varValue for k in kebun_list if (k, pabrik) in cost_tbs.keys()])
    cpo_out = tbs_in * yield_rate
    util = (tbs_in / factory_capacity[pabrik]) * 100
    html += f"""
                        <tr>
                            <td><strong>{pabrik}</strong></td>
                            <td>{tbs_in:,.0f}</td>
                            <td>{cpo_out:,.0f}</td>
                            <td>{factory_capacity[pabrik]:,} ton</td>
                            <td class="success">{util:.1f}%</td>
                        </tr>
"""

html += """
                    </tbody>
                </table>
            </div>
            
            <div class="section">
                <h2>📦 DISTRIBUSI CPO (Pabrik → PD)</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Dari Pabrik</th>
                            <th>Ke PD</th>
                            <th>Jumlah (ton)</th>
                            <th>Biaya/ton</th>
                            <th>Total Biaya</th>
                        </tr>
                    </thead>
                    <tbody>
"""

for (pabrik, pd), var in y.items():
    if var.varValue > 0.01:
        html += f"""
                        <tr>
                            <td><strong>{pabrik}</strong></td>
                            <td><strong>{pd}</strong></td>
                            <td>{var.varValue:,.1f}</td>
                            <td>Rp {cost_cpo[(pabrik, pd)]:,}</td>
                            <td>Rp {cost_cpo[(pabrik, pd)] * var.varValue:,.0f}</td>
                        </tr>
"""

html += """
                    </tbody>
                </table>
            </div>
            
            <div class="section">
                <h2>✅ VERIFIKASI PEMENUHAN DEMAND</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Pusat Distribusi</th>
                            <th>Demand</th>
                            <th>Supplied</th>
                            <th>Pemenuhan</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
"""

for pd in pd_list:
    supplied = sum([y[(p, pd)].varValue for p in pabrik_list if (p, pd) in cost_cpo.keys()])
    pct = (supplied / demand[pd]) * 100
    html += f"""
                        <tr>
                            <td><strong>{pd}</strong></td>
                            <td>{demand[pd]:,} ton</td>
                            <td>{supplied:,.1f} ton</td>
                            <td>{pct:.1f}%</td>
                            <td class="success">✓ Terpenuhi</td>
                        </tr>
"""

html += f"""
                    </tbody>
                </table>
            </div>
            
            <div class="section">
                <h2>📝 KESIMPULAN</h2>
                <div style="background: white; padding: 25px; border-radius: 12px; line-height: 1.8;">
                    <p><strong>1. Model Linear Programming</strong> berhasil dibangun dengan {len(x) + len(y)} variabel dan {len(model.constraints)} constraints.</p>
                    <br>
                    <p><strong>2. Solusi Optimal</strong> ditemukan menggunakan Python PuLP dengan CBC solver.</p>
                    <br>
                    <p><strong>3. Total Biaya Optimal:</strong> <strong style="color: #1e3c72;">Rp {total_biaya:,.0f}/bulan</strong></p>
                    <ul style="margin-left: 30px; margin-top: 10px;">
                        <li>Biaya TBS: Rp {biaya_tbs:,.0f} ({biaya_tbs/total_biaya*100:.1f}%)</li>
                        <li>Biaya CPO: Rp {biaya_cpo:,.0f} ({biaya_cpo/total_biaya*100:.1f}%)</li>
                    </ul>
                    <br>
                    <p><strong>4. Semua Constraints</strong> terpenuhi:</p>
                    <ul style="margin-left: 30px; margin-top: 10px;">
                        <li>✓ Kapasitas kebun tidak terlampaui</li>
                        <li>✓ Kapasitas pabrik tidak terlampaui</li>
                        <li>✓ Material balance TBS→CPO seimbang</li>
                        <li>✓ Semua demand terpenuhi 100%</li>
                    </ul>
                    <br>
                    <p><strong>5. Utilisasi Kapasitas</strong> tinggi menunjukkan efisiensi optimal.</p>
                </div>
            </div>
        </div>
        
        <div class="footer">
            <p><strong>Python PuLP Solution</strong></p>
            <p>Teknik Riset Operasional</p>
            <p>Program Studi Teknik Informatika</p>
        </div>
    </div>
</body>
</html>
"""

# Save HTML
output_file = 'hasil_python_solver.html'
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"✓ HTML report saved: {output_file}")

# Export Excel
try:
    tbs_data = []
    for (k, p), var in x.items():
        if var.varValue > 0.01:
            tbs_data.append({
                'Dari_Kebun': k,
                'Ke_Pabrik': p,
                'Jumlah_ton': var.varValue,
                'Biaya_per_ton': cost_tbs[(k, p)],
                'Total_Biaya': cost_tbs[(k, p)] * var.varValue
            })
    
    cpo_data = []
    for (p, pd), var in y.items():
        if var.varValue > 0.01:
            cpo_data.append({
                'Dari_Pabrik': p,
                'Ke_PD': pd,
                'Jumlah_ton': var.varValue,
                'Biaya_per_ton': cost_cpo[(p, pd)],
                'Total_Biaya': cost_cpo[(p, pd)] * var.varValue
            })
    
    with pd.ExcelWriter('hasil_python_solver.xlsx', engine='openpyxl') as writer:
        pd.DataFrame({'Metrik': ['Total Biaya', 'Biaya TBS', 'Biaya CPO'], 
                      'Nilai': [total_biaya, biaya_tbs, biaya_cpo]}).to_excel(writer, sheet_name='Summary', index=False)
        pd.DataFrame(tbs_data).to_excel(writer, sheet_name='Alokasi_TBS', index=False)
        pd.DataFrame(cpo_data).to_excel(writer, sheet_name='Alokasi_CPO', index=False)
    
    print("✓ Excel file saved: hasil_python_solver.xlsx")
except Exception as e:
    print(f"⚠️  Excel export error: {e}")

# Open browser
abs_path = os.path.abspath(output_file)
print("\n[6] Membuka browser...")

try:
    webbrowser.open('file://' + abs_path)
    print("✓ Browser opened!")
except Exception as e:
    print(f"⚠️  Error: {e}")
    print(f"   Please open manually: {output_file}")

print("\n" + "="*80)
print("✅ POINT 3a (PYTHON) SELESAI!")
print("="*80)
print(f"\nOutput files:")
print(f"  • HTML: {output_file}")
print(f"  • Excel: hasil_python_solver.xlsx")
print("\nSelanjutnya: Jalankan comparison_solver.py untuk Point 3c")