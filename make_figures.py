"""Render two portfolio figures from executed outputs; optionally preview XLSX sheets."""
import argparse
from datetime import datetime
import json
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import pandas as pd


def table_figure(title, columns, rows, path, size=(12,4)):
    fig, ax = plt.subplots(figsize=size)
    ax.axis('off')
    ax.set_title(title, fontsize=15, fontweight='bold', pad=20)
    table = ax.table(cellText=rows, colLabels=columns, loc='center', cellLoc='left')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1,1.7)
    for (row,col), cell in table.get_celld().items():
        cell.set_edgecolor('white')
        if row == 0:
            cell.set_facecolor('#24364b')
            cell.set_text_props(color='white', weight='bold')
        else:
            cell.set_facecolor('#edf2f6' if row % 2 else '#f8fafc')
    fig.tight_layout()
    fig.savefig(path,dpi=160,facecolor='white')
    plt.close(fig)


def generate(output, portfolio, qa_dir=None):
    output, portfolio = Path(output), Path(portfolio)
    portfolio.mkdir(parents=True,exist_ok=True)
    report = json.loads((output/'report.json').read_text())
    audit = pd.read_csv(output/'before_after.csv',dtype=str,keep_default_na=False)
    examples = audit.iloc[[0,1,10,14]]
    cols = ['source_row','raw_category','clean_category','raw_order_date','clean_order_date','decision']
    table_figure('Automated Excel & CSV cleaning: actual before / after', ['Source row','Raw category','Clean category','Raw date','Clean date','Disposition'],examples[cols].values.tolist(),portfolio/'before-after.png')
    fig, axes = plt.subplots(1,2,figsize=(10,4.6))
    labels = ['Input','Accepted','Duplicates','Rejected']
    counts = [report[k] for k in ('input_rows','cleaned_rows','duplicates_removed','rejected_rows')]
    axes[0].bar(labels,counts,color=['#8295ab','#298b78','#dda447','#c96565'])
    for x,y in enumerate(counts):
        axes[0].text(x,y+.3,str(y),ha='center')
    axes[0].set_ylim(0,max(counts)*1.2)
    axes[0].set_ylabel('Records')
    axes[0].set_title('Every input row accounted for')
    axes[1].bar(list(report['category_counts']),list(report['category_counts'].values()),color='#426a99')
    for x,y in enumerate(report['category_counts'].values()):
        axes[1].text(x,y+.1,str(y),ha='center')
    axes[1].set_ylim(0,max(report['category_counts'].values(),default=1)*1.3)
    axes[1].set_title('Accepted records by category')
    axes[1].set_ylabel('Records')
    fig.suptitle('Automated Excel & CSV Data Cleaning Pipeline',fontsize=14,fontweight='bold')
    fig.text(.5,.02,f"Measured results: {report['normalized_cells']} normalized cells | USD {report['amount_total']} accepted amount total",ha='center',fontsize=10)
    fig.tight_layout(rect=(0,.065,1,.94))
    fig.savefig(portfolio/'cleaning-summary.png',dpi=160,facecolor='white')
    plt.close(fig)
    if qa_dir:
        qa_dir = Path(qa_dir)
        qa_dir.mkdir(parents=True,exist_ok=True)
        wb = load_workbook(output/'cleaned.xlsx',data_only=True)
        for sheet in wb:
            rows = []
            for cells in sheet:
                values = []
                for cell in cells:
                    v = cell.value
                    if isinstance(v,datetime):
                        v = v.strftime('%Y-%m-%d')
                    elif isinstance(v,(int,float)) and '$' in cell.number_format:
                        v = f'${v:,.2f}'
                    values.append('' if v is None else str(v))
                rows.append(values)
            table_figure('Exported workbook: '+sheet.title,rows[0],rows[1:],qa_dir/(sheet.title+'.png'),size=(15,7))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output',type=Path,default=Path(__file__).parent/'sample_output')
    parser.add_argument('--portfolio',type=Path,default=Path(__file__).parent/'portfolio')
    parser.add_argument('--qa-dir',type=Path)
    args = parser.parse_args()
    generate(args.output,args.portfolio,args.qa_dir)
