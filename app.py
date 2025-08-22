from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, send_file
from werkzeug.utils import secure_filename
import os
import datetime as dt
import pandas as pd
from io import BytesIO

from config import Config

app = Flask(__name__)
app.config.from_object(Config)

ALLOWED_EXTENSIONS = {'.xlsx', '.xls'}

REQUIRED_FILES = {
    'corp_mid_income_detail': '对公业务-分行中间业务收入统计客户收入明细【即期、远期、掉期、期权】',
    'retail_spot_income': '对私业务-即期收入',
    'irs_income_detail': '利率互换收入明细',
    'cust_branch_map': '客户和管户机构（支行）对应关系（客户名称、所属中心支行）',
    'org_code_map': '机构号映射表（名称、分支机构号）'
}

TEMPLATE_COLUMNS = {
    'corp_mid_income_detail': ['客户名称','即期','远期','掉期','期权'],
    'retail_spot_income': ['客户名称','机构号','损益金额'],
    'irs_income_detail': ['客户名称','分行落账损益'],
    'cust_branch_map': ['客户名称','所属中心支行'],
    'org_code_map': ['名称','分支机构号']
}

NUM_COLS_A = ['即期','远期','掉期','期权']
REQ_COLS = {
    'corp_mid_income_detail': ['客户名称'] + NUM_COLS_A,
    'cust_branch_map': ['客户名称','所属中心支行'],
    'irs_income_detail': ['客户名称','分行落账损益'],
    'retail_spot_income': ['客户名称','机构号','损益金额'],
    'org_code_map': ['名称','分支机构号']
}

def allowed_file(filename: str) -> bool:
    ext = os.path.splitext(filename)[1].lower()
    return ext in ALLOWED_EXTENSIONS

def to_numeric(series):
    return pd.to_numeric(series, errors='coerce').fillna(0.0)

@app.route('/template/<key>')
def template_download(key):
    if key not in TEMPLATE_COLUMNS:
        flash('未知的模板键。', 'danger')
        return redirect(url_for('index'))
    cols = TEMPLATE_COLUMNS[key]
    examples = {
        'corp_mid_income_detail': [
            {'客户名称':'示例客户A','即期':1000,'远期':0,'掉期':0,'期权':0},
            {'客户名称':'示例客户B','即期':0,'远期':200,'掉期':50,'期权':10}
        ],
        'retail_spot_income': [
            {'客户名称':'示例客户C','机构号':'123456','损益金额':300},
            {'客户名称':'示例客户D','机构号':'654321','损益金额':-50}
        ],
        'irs_income_detail': [
            {'客户名称':'示例客户E','分行落账损益':120},
            {'客户名称':'示例客户F','分行落账损益':-30}
        ],
        'cust_branch_map': [
            {'客户名称':'示例客户A','所属中心支行':'XX中心支行'},
            {'客户名称':'示例客户B','所属中心支行':'YY中心支行'}
        ],
        'org_code_map': [
            {'名称':'AA支行','分支机构号':'100001'},
            {'名称':'BB支行','分支机构号':'100002'}
        ]
    }.get(key, [])
    df = pd.DataFrame(examples, columns=cols)
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='模板')
    bio.seek(0)
    filename = f"{key}_模板.xlsx"
    return send_file(bio, as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/download/<path:ts>/<path:filename>')
def download(ts, filename):
    base = app.config['UPLOAD_FOLDER']
    folder = os.path.join(base, ts)
    return send_from_directory(folder, filename, as_attachment=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        missing = [desc for key, desc in REQUIRED_FILES.items()
                   if key not in request.files or request.files[key].filename.strip() == '']
        if missing:
            flash(f"缺少文件：{'、'.join(missing)}", 'danger')
            return redirect(url_for('index'))

        ts = dt.datetime.now().strftime('%Y%m%d_%H%M%S')
        updir = os.path.join(app.config['UPLOAD_FOLDER'], ts)
        os.makedirs(updir, exist_ok=True)

        dfs = {}
        summaries = {}
        for field, desc in REQUIRED_FILES.items():
            f = request.files[field]
            filename = secure_filename(f.filename)
            if not allowed_file(filename):
                flash(f"{desc} 仅支持 .xlsx/.xls 文件：{filename}", 'danger')
                return redirect(url_for('index'))
            save_path = os.path.join(updir, f"{field}_{filename}")
            f.save(save_path)

            try:
                df = pd.read_excel(save_path, sheet_name=0)
            except Exception as e:
                flash(f"读取 {desc} 失败：{e}", 'danger')
                return redirect(url_for('index'))
            need = REQ_COLS.get(field, [])
            missing_cols = [c for c in need if c not in df.columns]
            if missing_cols:
                flash(f"{desc} 缺少列：{', '.join(missing_cols)}", 'danger')
                return redirect(url_for('index'))

            dfs[field] = df
            summaries[field] = {
                'desc': desc,
                'filename': filename,
                'rows': int(df.shape[0]),
                'cols': int(df.shape[1]),
                'columns': [str(c) for c in df.columns[:12]]
            }

        # A
        corp = dfs['corp_mid_income_detail'].copy()
        cust_map = dfs['cust_branch_map'].copy()
        for c in NUM_COLS_A:
            corp[c] = to_numeric(corp[c])
        corp_g = corp.groupby('客户名称', as_index=False)[NUM_COLS_A].sum()
        cust_map_sub = cust_map[['客户名称','所属中心支行']].drop_duplicates()
        A = corp_g.merge(cust_map_sub, on='客户名称', how='left')
        A = A[['客户名称','所属中心支行'] + NUM_COLS_A]
        A.to_excel(os.path.join(updir, 'A_对公聚合.xlsx'), index=False)

        # B
        irs = dfs['irs_income_detail'].copy()
        if irs.shape[0] > 0:
            irs = irs.iloc[1:, :].reset_index(drop=True)
        irs['分行落账损益'] = to_numeric(irs['分行落账损益'])
        irs_g = irs.groupby('客户名称', as_index=False)['分行落账损益'].sum().rename(columns={'分行落账损益':'利率互换'})
        B = irs_g.merge(cust_map_sub, on='客户名称', how='left')
        B = B[['客户名称','所属中心支行','利率互换']]
        B.to_excel(os.path.join(updir, 'B_利率互换聚合.xlsx'), index=False)

        # C
        C = A.merge(B[['客户名称','所属中心支行','利率互换']], on='客户名称', how='outer', suffixes=('_A','_B'))
        C['所属中心支行'] = C['所属中心支行_A'].fillna(C['所属中心支行_B'])
        C.drop(columns=['所属中心支行_A','所属中心支行_B'], inplace=True)
        for col in NUM_COLS_A + ['利率互换']:
            if col not in C.columns:
                C[col] = 0.0
            C[col] = to_numeric(C[col])
        C = C[['客户名称','所属中心支行'] + NUM_COLS_A + ['利率互换']]
        C.to_excel(os.path.join(updir, 'C_汇总.xlsx'), index=False)

        # D
        retail = dfs['retail_spot_income'].copy()
        orgmap = dfs['org_code_map'].copy()
        retail['损益金额'] = to_numeric(retail['损益金额'])
        orgmap_sub = orgmap[['名称','分支机构号']].drop_duplicates().rename(columns={'分支机构号':'机构号'})
        D = retail.merge(orgmap_sub, on='机构号', how='left')
        D = D[['客户名称','机构号','名称','损益金额']]
        D.to_excel(os.path.join(updir, 'D_对私映射.xlsx'), index=False)

        # E
        D_as_C = pd.DataFrame({
            '客户名称': D['客户名称'],
            '所属中心支行': D['名称'],
            '即期': D['损益金额'],
            '远期': 0.0,
            '掉期': 0.0,
            '期权': 0.0,
            '利率互换': 0.0
        })
        for col in ['即期','远期','掉期','期权','利率互换']:
            D_as_C[col] = to_numeric(D_as_C[col])
        E = pd.concat([C, D_as_C], ignore_index=True)
        E.to_excel(os.path.join(updir, 'E_合并结果.xlsx'), index=False)

        pivot = E.groupby('所属中心支行')[['即期','远期','掉期','期权','利率互换']].sum().reset_index()
        pivot['合计'] = pivot[['即期','远期','掉期','期权','利率互换']].sum(axis=1)
        pivot.to_excel(os.path.join(updir, 'Pivot_按支行汇总.xlsx'), index=False)

        downloads = [
            ('A_对公聚合.xlsx', f'/download/{ts}/A_对公聚合.xlsx'),
            ('B_利率互换聚合.xlsx', f'/download/{ts}/B_利率互换聚合.xlsx'),
            ('C_汇总.xlsx', f'/download/{ts}/C_汇总.xlsx'),
            ('D_对私映射.xlsx', f'/download/{ts}/D_对私映射.xlsx'),
            ('E_合并结果.xlsx', f'/download/{ts}/E_合并结果.xlsx'),
            ('Pivot_按支行汇总.xlsx', f'/download/{ts}/Pivot_按支行汇总.xlsx'),
        ]

        return render_template('result.html',
                               summaries=summaries,
                               upload_dir_rel=os.path.relpath(updir, start=os.getcwd()),
                               downloads=downloads)

    return render_template('index.html', required_files=REQUIRED_FILES, templates_keys=list(TEMPLATE_COLUMNS.keys()))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
