import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import scipy.stats as stats
from statsmodels.stats.diagnostic import normal_ad
import io
import re

st.set_page_config(page_title="스텐트 성능 분석", layout="wide")
st.title("📊 스텐트 자동 분석 대시보드 (스마트 판정 및 알람)")

pink_fill = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")

# ---------------------------------------------------------
# 성능 시험용 - 부등호 및 양측범위(~) 판독 함수
# ---------------------------------------------------------
def parse_criteria(crit_str):
    if crit_str is None: return "<=", 0.0
    crit_str = str(crit_str).replace(',', '').strip()
    
    if '~' in crit_str:
        parts = crit_str.split('~')
        if len(parts) == 2:
            try:
                return "~", (float(parts[0].strip()), float(parts[1].strip()))
            except ValueError:
                pass
                
    match = re.match(r"([<>]=?)\s*([\d\.]+)", crit_str)
    if match:
        return match.group(1), float(match.group(2))
    
    try:
        return "<=", float(crit_str)
    except ValueError:
        return "<=", 0.0

def check_data_pass(value, operator, limit):
    if operator == '~': return limit[0] <= value <= limit[1]
    if operator == '<=': return value <= limit
    if operator == '<': return value < limit
    if operator == '>=': return value >= limit
    if operator == '>': return value > limit
    return True

# ---------------------------------------------------------
# 미니탭 스타일 그래프 생성 함수
# ---------------------------------------------------------
def create_minitab_plot(data, model_name):
    n = len(data)
    mean = np.mean(data)
    std = np.std(data, ddof=1)
    var = np.var(data, ddof=1)
    skew = stats.skew(data, bias=False)
    kurt = stats.kurtosis(data, bias=False)
    stat_ad, p_val_ad = normal_ad(data)
    minimum = np.min(data)
    q1 = np.percentile(data, 25)
    median = np.median(data)
    q3 = np.percentile(data, 75)
    maximum = np.max(data)
    ci_mean = stats.t.interval(0.95, df=n-1, loc=mean, scale=std/np.sqrt(n))
    ci_median = (median - 1.57*(q3-q1)/np.sqrt(n), median + 1.57*(q3-q1)/np.sqrt(n))
    ci_std = (std * np.sqrt((n-1)/stats.chi2.ppf(0.975, n-1)), std * np.sqrt((n-1)/stats.chi2.ppf(0.025, n-1)))

    fig = plt.figure(figsize=(10, 6))
    gs = gridspec.GridSpec(3, 2, width_ratios=[2.5, 1], height_ratios=[3, 1, 1.2])
    fig.suptitle(f"Summary Report for {model_name}", fontsize=14)

    ax1 = fig.add_subplot(gs[0, 0])
    ax1.hist(data, bins=5, density=True, alpha=0.6, color='#5c8ccc', edgecolor='black')
    xmin, xmax = ax1.get_xlim()
    x = np.linspace(xmin, xmax, 100)
    p = stats.norm.pdf(x, mean, std)
    ax1.plot(x, p, 'darkred', linewidth=1.5)
    ax1.grid(True, alpha=0.3)
    ax1.set_yticks([])

    ax2 = fig.add_subplot(gs[1, 0], sharex=ax1)
    ax2.boxplot(data, vert=False, patch_artist=True, boxprops=dict(facecolor='#5c8ccc', color='black'))
    ax2.set_yticks([])
    ax2.grid(True, alpha=0.3)

    ax3 = fig.add_subplot(gs[2, 0], sharex=ax1)
    ax3.errorbar((ci_mean[0]+ci_mean[1])/2, 1, xerr=(ci_mean[1]-ci_mean[0])/2, fmt='o', color='#0055a4', capsize=4)
    ax3.errorbar((ci_median[0]+ci_median[1])/2, 0.5, xerr=(ci_median[1]-ci_median[0])/2, fmt='o', color='#0055a4', capsize=4)
    ax3.set_yticks([0.5, 1])
    ax3.set_yticklabels(['Median', 'Mean'])
    ax3.set_ylim(0, 1.5)
    ax3.set_title("95% Confidence Intervals", fontsize=10)
    ax3.grid(True, alpha=0.3)

    ax4 = fig.add_subplot(gs[:, 1])
    ax4.axis('off')
    stats_text = (
        f"Anderson-Darling Normality Test\n"
        f"{'A-Squared':<15} {stat_ad:.2f}\n"
        f"{'P-Value':<15} {p_val_ad:.3f}\n\n"
        f"{'Mean':<15} {mean:.1f}\n"
        f"{'StDev':<15} {std:.1f}\n"
        f"{'N':<15} {n}\n\n"
        f"{'Minimum':<15} {minimum:.1f}\n"
        f"{'1st Quartile':<15} {q1:.1f}\n"
        f"{'Median':<15} {median:.1f}\n"
        f"{'3rd Quartile':<15} {q3:.1f}\n"
        f"{'Maximum':<15} {maximum:.1f}\n"
    )
    ax4.text(0.1, 0.95, stats_text, va='top', ha='left', fontsize=9, family='monospace')
    plt.tight_layout()
    
    img_buffer = io.BytesIO()
    plt.savefig(img_buffer, format='png', dpi=150)
    img_buffer.seek(0)
    
    return fig, img_buffer

# ---------------------------------------------------------
# 웹 화면 및 메인 분석 로직
# ---------------------------------------------------------
uploaded_file = st.file_uploader("여러 시험 데이터가 세로로 나열된 엑셀 파일을 끌어다 놓으세요.", type=['xlsx'])

if uploaded_file is not None:
    wb = openpyxl.load_workbook(uploaded_file)
    ws = wb.active
    
    try:
        k_val_one = float(ws['D2'].value)
        k_val_two = float(ws['E2'].value)
    except:
        st.warning("⚠ D2(단측) 또는 E2(양측) 셀에 K값이 비어있을 수 있습니다. 성능 시험이 포함된 경우 에러가 발생할 수 있습니다.")
        k_val_one, k_val_two = 0.0, 0.0
        
    blocks = []
    for r in range(1, ws.max_row + 15):
        cell_val = ws.cell(row=r, column=1).value
        if isinstance(cell_val, str) and cell_val.strip() == "평균":
            header_row = r - 1
            while header_row > 1:
                h_val = ws.cell(row=header_row, column=1).value
                if h_val is not None:
                    h_str = str(h_val).strip()
                    if not h_str.replace('.', '', 1).isdigit():
                        break
                header_row -= 1
            blocks.append({'mean_row': r, 'header_row': header_row})
            
    if not blocks:
        st.error("⚠️ A열에 '평균'이 포함되어 있는지 확인하세요.")
        st.stop()
        
    st.success(f"✅ 총 **{len(blocks)}개** 세부 묶음 스캔 완료!")
    
    for idx, block in enumerate(blocks):
        mean_row = block['mean_row']
        header_row = block['header_row']
        test_name = str(ws.cell(row=header_row, column=1).value)
        
        last_col = 2
        while ws.cell(row=header_row, column=last_col).value is not None:
            last_col += 1
        last_col -= 1
        
        # [핵심] 치수 시험 vs 성능 시험 자동 판별 (A열에 K값 유무 확인)
        is_dimensional = True
        crit_row = None
        for r in range(mean_row, mean_row + 6):
            c_val = ws.cell(row=r, column=1).value
            if c_val is not None:
                c_str = str(c_val).replace(" ", "").lower()
                if "criteria" in c_str:
                    crit_row = r
                if "k값" in c_str or "p값" in c_str:
                    is_dimensional = False
                    
        if crit_row is None:
            crit_row = mean_row + 4 # 기본값

        mode_text = "📏 치수 시험 모드" if is_dimensional else "📈 성능 시험 모드"
        st.markdown(f"### 🧪 {idx+1}번 묶음: `{test_name}` ({mode_text})")
        cols = st.columns(max(1, last_col - 1))
        col_idx = 0
        img_insert_col = max(7, last_col + 2)
        
        for col in range(2, last_col + 1):
            model_name = ws.cell(row=header_row, column=col).value
            if not model_name: continue
                
            raw_criteria = ws.cell(row=crit_row, column=col).value
            
            # 원본 데이터 수집
            data_cells = []
            data = []
            for dr in range(header_row + 1, mean_row):
                d_val = ws.cell(row=dr, column=col).value
                if d_val is not None:
                    val_float = float(str(d_val).replace(',', ''))
                    data.append(val_float)
                    data_cells.append((dr, val_float))
                    
            if not data: continue
                
            mean_val = np.mean(data)
            std_val = np.std(data, ddof=1)
            
            if is_dimensional:
                # -----------------------------------------------------
                # 치수 시험 모드: 타겟값과 %허용오차를 계산하여 불량만 판별
                # -----------------------------------------------------
                try:
                    base_val = float(str(raw_criteria).replace(',', ''))
                except:
                    base_val = 0.0
                    
                pct_val = ws.cell(row=crit_row+1, column=col).value
                pct_float = 0.0
                if isinstance(pct_val, str) and '%' in pct_val:
                    pct_float = float(pct_val.replace('%', '')) / 100.0
                elif isinstance(pct_val, (int, float)):
                    if pct_val > 1: pct_float = pct_val / 100.0
                    else: pct_float = float(pct_val)
                    
                lower_limit = base_val * (1 - pct_float)
                upper_limit = base_val * (1 + pct_float)
                
                failed_count = 0
                for dr, val_float in data_cells:
                    if not (lower_limit <= val_float <= upper_limit):
                        ws.cell(row=dr, column=col).fill = pink_fill
                        failed_count += 1
                        
                ws.cell(row=mean_row, column=col).value = round(mean_val, 2)
                ws.cell(row=mean_row+1, column=col).value = round(std_val, 2)
                
                with cols[col_idx % len(cols)]:
                    st.write(f"**{model_name}**")
                    limit_text = f"기준: {round(lower_limit, 2)} ~ {round(upper_limit, 2)}"
                    if failed_count > 0:
                        st.error(f"❌ 불량 {failed_count}건 발생\n\n{limit_text}")
                    else:
                        st.success(f"✅ 전수 통과\n\n{limit_text}")
                
            else:
                # -----------------------------------------------------
                # 성능 시험 모드: UTL/LTL 계산 및 미니탭 그래프 출력
                # -----------------------------------------------------
                operator, limit = parse_criteria(raw_criteria)
                
                for dr, val_float in data_cells:
                    if not check_data_pass(val_float, operator, limit):
                        ws.cell(row=dr, column=col).fill = pink_fill
                        
                stat, p_val = normal_ad(np.array(data))
                
                if operator == '~':
                    ltl = mean_val - (k_val_two * std_val)
                    utl = mean_val + (k_val_two * std_val)
                    bound_text = f"{round(ltl, 2)} ~ {round(utl, 2)}"
                    is_pass = (ltl >= limit[0]) and (utl <= limit[1])
                    limit_text = f"{limit[0]} ~ {limit[1]}"
                else:
                    if operator in ['<=', '<']:
                        bound_val = mean_val + (k_val_one * std_val)
                    else:
                        bound_val = mean_val - (k_val_one * std_val)
                    bound_text = round(bound_val, 2)
                    is_pass = check_data_pass(bound_val, operator, limit)
                    limit_text = f"{operator} {limit}"
                
                ws.cell(row=mean_row, column=col).value = round(mean_val, 2)
                ws.cell(row=mean_row+1, column=col).value = round(std_val, 2)
                
                cell_k = ws.cell(row=mean_row+2, column=col)
                cell_k.value = bound_text
                if not is_pass: cell_k.fill = pink_fill
                    
                cell_p = ws.cell(row=mean_row+3, column=col)
                cell_p.value = round(p_val, 3)
                if p_val <= 0.05: cell_p.fill = pink_fill
                        
                fig, img_buffer = create_minitab_plot(np.array(data), model_name)
                img = OpenpyxlImage(img_buffer)
                img.width = int(9.6 * 37.8)
                img.height = int(6.2 * 37.8)
                
                c_letter = get_column_letter(img_insert_col)
                ws.add_image(img, f"{c_letter}{header_row}")
                img_insert_col += 6
                
                with cols[col_idx % len(cols)]:
                    st.write(f"**{model_name}**")
                    if is_pass: st.success(f"✅ 통과 ({bound_text} | 기준: {limit_text})")
                    else: st.error(f"❌ 불합격 ({bound_text} 🚫 기준: {limit_text})")
                    st.pyplot(fig)
                plt.close(fig)
                
            col_idx += 1
            
        st.divider()

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)

    st.subheader("🎉 통합 분석 완료!")
    st.download_button(
        label="📥 결과 엑셀 다운로드 (스마트 판정 적용)",
        data=excel_buffer,
        file_name="스텐트_전체성능분석_완료.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
