"""
SMT 개발 모델 PL 통합 관리 대시보드
Streamlit Community Cloud 배포용
"""

import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import json

# ===== 페이지 설정 =====
st.set_page_config(
    page_title="SMT 개발 모델 관리",
    page_icon="📋",
    layout="wide"
)

# ===== 세션 상태 초기화 =====
if "data" not in st.session_state:
    st.session_state.data = {
        "models": [],
        "milestones": [],
        "defects": [],
        "bom_changes": [],
        "fa_records": []
    }

data = st.session_state.data

# ===== 공통 함수 =====
def get_next_id(prefix, items):
    if not items:
        return f"{prefix}001"
    nums = []
    for item in items:
        try:
            nums.append(int(item["id"].replace(prefix, "")))
        except:
            nums.append(0)
    return f"{prefix}{max(nums)+1:03d}"

# ===== 사이드바 =====
with st.sidebar:
    st.title("📋 SMT PL 관리")
    st.caption("v2.0 | Cloud Edition")
    st.divider()

    # ===== 모델 등록 =====
    st.subheader("➕ 모델 등록")
    with st.form("add_model", clear_on_submit=True):
        m_name = st.text_input("모델명", placeholder="예: iPhone 18 Pro PBA")
        m_customer = st.selectbox("고객사", ["Apple", "Samsung", "기타"])
        m_col1, m_col2 = st.columns(2)
        with m_col1:
            m_type = st.selectbox("제품유형", ["스마트폰", "태블릿", "노트북", "모니터", "TV", "웨어러블", "기타"])
        with m_col2:
            m_phase = st.selectbox("개발단계", ["Proto", "EVT", "DVT", "PVT", "MP"])
        m_start = st.date_input("시작일", value=date.today())
        m_target = st.date_input("목표완료일", value=date.today() + timedelta(days=90))
        m_pl = st.text_input("PL 담당자", placeholder="이름 입력")

        if st.form_submit_button("모델 등록", use_container_width=True):
            if m_name:
                data["models"].append({
                    "id": get_next_id("M", data["models"]),
                    "name": m_name,
                    "customer": m_customer,
                    "type": m_type,
                    "phase": m_phase,
                    "status": "진행중",
                    "start": str(m_start),
                    "target": str(m_target),
                    "pl": m_pl,
                    "issue": "",
                    "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "modified_by": m_pl,
                    "modified_at": datetime.now().strftime("%Y-%m-%d %H:%M")
                })
                st.success(f"✅ {m_name} 등록!")
                st.rerun()
            else:
                st.error("모델명 입력 필요!")

    st.divider()

    # 등록 모델 목록
    if data["models"]:
        st.subheader("📱 등록 모델")
        for m in data["models"]:
            icon = {"진행중":"🟢","지연":"🟡","이슈":"🔴","완료":"✅","보류":"⚪"}.get(m["status"],"⚪")
            st.caption(f"{icon} {m['name']} ({m['phase']})")

    st.divider()

    # 데이터 내보내기/가져오기
    st.subheader("💾 데이터 관리")

    # JSON 내보내기
    if data["models"]:
        json_str = json.dumps(data, ensure_ascii=False, indent=2, default=str)
        st.download_button("📥 데이터 백업 (JSON)", json_str,
                          file_name=f"pl_backup_{datetime.now().strftime('%Y%m%d')}.json",
                          mime="application/json", use_container_width=True)

    # JSON 가져오기
    uploaded = st.file_uploader("📤 데이터 복원 (JSON)", type="json")
    if uploaded:
        try:
            imported = json.loads(uploaded.read().decode("utf-8"))
            st.session_state.data = imported
            st.success("✅ 데이터 복원 완료!")
            st.rerun()
        except:
            st.error("파일 형식 오류!")

    # 엑셀 내보내기
    if data["models"]:
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for key, name in [("models","모델현황"),("milestones","마일스톤"),
                             ("defects","불량추적"),("bom_changes","BOM변경"),("fa_records","FA기록")]:
                if data[key]:
                    pd.DataFrame(data[key]).to_excel(writer, sheet_name=name, index=False)
        output.seek(0)
        st.download_button("📥 엑셀 내보내기", output,
                          file_name=f"PL관리_{datetime.now().strftime('%Y%m%d')}.xlsx",
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                          use_container_width=True)

    if st.button("⚠️ 전체 초기화", use_container_width=True):
        st.session_state.data = {"models":[],"milestones":[],"defects":[],"bom_changes":[],"fa_records":[]}
        st.rerun()

# ===== 메인 영역 =====
if not data["models"]:
    st.title("📋 SMT 개발 모델 PL 관리 대시보드")
    st.info("👈 왼쪽 사이드바에서 개발 모델을 먼저 등록해주세요!")
    st.stop()

model_names = [m["name"] for m in data["models"]]

# ===== 탭 구성 =====
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 대시보드", "📅 마일스톤", "❌ 불량 추적", "📦 BOM 변경", "🔬 FA 기록", "⚙️ 모델 관리"
])

# ========== 탭1: 대시보드 ==========
with tab1:
    st.header("📊 개발 모델 현황")

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric("전체 모델", f"{len(data['models'])}건")
    with c2:
        st.metric("진행중", f"{len([m for m in data['models'] if m['status']=='진행중'])}건")
    with c3:
        st.metric("이슈/지연", f"{len([m for m in data['models'] if m['status'] in ['이슈','지연']])}건")
    with c4:
        st.metric("불량 건수", f"{len(data['defects'])}건")
    with c5:
        open_fa = len([f for f in data["fa_records"] if f.get("status") != "완료"])
        st.metric("미완료 FA", f"{open_fa}건")

    st.divider()

    for m in data["models"]:
        status_icon = {"진행중":"🟢","지연":"🟡","이슈":"🔴","완료":"✅","보류":"⚪"}.get(m["status"],"⚪")
        with st.expander(f"{status_icon} {m['name']}  |  {m['customer']}  |  {m['phase']}  |  PL: {m.get('pl','-')}", expanded=False):
            ec1, ec2, ec3, ec4 = st.columns(4)
            with ec1:
                st.caption("고객사 / 유형")
                st.write(f"**{m['customer']}** / {m['type']}")
            with ec2:
                st.caption("개발 단계")
                st.write(f"**{m['phase']}**")
            with ec3:
                st.caption("상태")
                st.write(f"**{m['status']}**")
            with ec4:
                target = datetime.strptime(m["target"], "%Y-%m-%d").date()
                days = (target - date.today()).days
                if days > 0:
                    st.caption(f"D-Day")
                    st.write(f"**D-{days}**")
                elif days == 0:
                    st.warning("⚠️ 오늘!")
                else:
                    st.error(f"🚨 D+{abs(days)} 초과")

            if m.get("issue"):
                st.info(f"📌 주요 이슈: {m['issue']}")

            mc1, mc2, mc3 = st.columns(3)
            with mc1:
                model_defects = len([d for d in data["defects"] if d["model"] == m["name"]])
                st.caption(f"불량 {model_defects}건")
            with mc2:
                model_fa = len([f for f in data["fa_records"] if f["model"] == m["name"]])
                st.caption(f"FA {model_fa}건")
            with mc3:
                model_bom = len([b for b in data["bom_changes"] if b["model"] == m["name"]])
                st.caption(f"BOM변경 {model_bom}건")

    # 불량 유형 차트
    if data["defects"]:
        st.divider()
        st.subheader("📈 전체 불량 현황")
        ch1, ch2 = st.columns(2)
        with ch1:
            type_counts = {}
            for d in data["defects"]:
                t = d["type"]
                type_counts[t] = type_counts.get(t, 0) + d.get("count", 1)
            if type_counts:
                st.write("**불량유형별 집계**")
                st.bar_chart(pd.DataFrame({"수량": type_counts}, index=type_counts.keys()))
        with ch2:
            proc_counts = {}
            for d in data["defects"]:
                p2 = d["process"]
                proc_counts[p2] = proc_counts.get(p2, 0) + d.get("count", 1)
            if proc_counts:
                st.write("**공정별 집계**")
                st.bar_chart(pd.DataFrame({"수량": proc_counts}, index=proc_counts.keys()))


# ========== 탭2: 마일스톤 ==========
with tab2:
    st.header("📅 마일스톤 관리")

    sel_model_ms = st.selectbox("모델 선택", model_names, key="ms_model")

    with st.form("add_ms", clear_on_submit=True):
        ms_c1, ms_c2, ms_c3 = st.columns([3, 1.5, 1.5])
        with ms_c1:
            ms_name = st.text_input("마일스톤명", placeholder="예: 리플로우 프로파일 최적화")
        with ms_c2:
            ms_date = st.date_input("목표일", value=date.today() + timedelta(days=14), key="ms_date")
        with ms_c3:
            ms_pri = st.selectbox("우선순위", ["높음", "중간", "낮음"])

        ms_r1, ms_r2 = st.columns(2)
        with ms_r1:
            ms_person = st.text_input("담당자", placeholder="이름", key="ms_person")
        with ms_r2:
            ms_note = st.text_input("비고", placeholder="관련 이슈 등", key="ms_note")

        if st.form_submit_button("➕ 마일스톤 추가", use_container_width=True):
            if ms_name:
                data["milestones"].append({
                    "id": get_next_id("MS", data["milestones"]),
                    "model": sel_model_ms,
                    "name": ms_name,
                    "date": str(ms_date),
                    "priority": ms_pri,
                    "status": "대기",
                    "person": ms_person,
                    "note": ms_note,
                    "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "modified_by": ms_person,
                    "modified_at": datetime.now().strftime("%Y-%m-%d %H:%M")
                })
                st.success("✅ 추가 완료!")
                st.rerun()

    model_ms = [ms for ms in data["milestones"] if ms["model"] == sel_model_ms]
    if model_ms:
        st.divider()
        for i, ms in enumerate(model_ms):
            pri_icon = {"높음":"🔴","중간":"🟡","낮음":"🟢"}.get(ms["priority"],"⚪")
            stat_icon = {"대기":"⬜","진행중":"🔵","완료":"✅","지연":"🟠","취소":"⛔"}.get(ms["status"],"⬜")

            col1, col2, col3, col4, col5 = st.columns([3, 1.5, 1, 1, 0.5])
            with col1:
                st.write(f"{pri_icon} {stat_icon} **{ms['name']}**")
                if ms.get("note"):
                    st.caption(f"📝 {ms['note']}  |  👤 {ms.get('person','-')}")
            with col2:
                target_d = datetime.strptime(ms["date"], "%Y-%m-%d").date()
                days = (target_d - date.today()).days
                if days > 0:
                    st.write(f"📆 {ms['date']} (D-{days})")
                elif days == 0:
                    st.warning("오늘!")
                else:
                    st.error(f"D+{abs(days)} 초과")
            with col3:
                new_st = st.selectbox("상태", ["대기","진행중","완료","지연","취소"],
                    index=["대기","진행중","완료","지연","취소"].index(ms["status"]),
                    key=f"ms_st_{ms['id']}", label_visibility="collapsed")
                if new_st != ms["status"]:
                    ms["status"] = new_st
                    ms["modified_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            with col4:
                st.caption(f"수정: {ms.get('modified_at','')[:10]}")
            with col5:
                if st.button("🗑️", key=f"del_ms_{ms['id']}"):
                    data["milestones"].remove(ms)
                    st.rerun()
    else:
        st.info("등록된 마일스톤이 없습니다.")


# ========== 탭3: 불량 추적 ==========
with tab3:
    st.header("❌ 불량 추적")

    sel_model_df = st.selectbox("모델 선택", model_names, key="df_model")

    with st.form("add_defect", clear_on_submit=True):
        df_r1c1, df_r1c2, df_r1c3 = st.columns(3)
        with df_r1c1:
            df_process = st.selectbox("공정", [
                "Solder Paste Print", "Pick & Place", "Reflow",
                "AOI", "X-ray", "ICT", "FCT", "Underfill", "Wire Bond", "기타"
            ])
        with df_r1c2:
            df_type = st.selectbox("불량유형", [
                "솔더 브릿지", "미실장", "위치 틀어짐", "냉납", "솔더볼",
                "리프트", "크랙", "위킹", "보이드", "인쇄불량", "기타"
            ])
        with df_r1c3:
            df_count = st.number_input("불량수량", min_value=1, value=1, key="df_count")

        df_r2c1, df_r2c2, df_r2c3 = st.columns(3)
        with df_r2c1:
            df_total = st.number_input("총생산수량", min_value=1, value=10000, key="df_total")
        with df_r2c2:
            df_loc = st.text_input("불량위치", placeholder="예: U23 BGA, R105")
        with df_r2c3:
            df_lot = st.text_input("LOT번호", placeholder="예: LOT-2026-0404")

        df_desc = st.text_area("상세내용", placeholder="불량 설명, 발생 조건 등", height=80)
        df_person = st.text_input("등록자", placeholder="이름", key="df_person")

        if st.form_submit_button("➕ 불량 등록", use_container_width=True, type="primary"):
            rate = round((df_count / df_total) * 100, 4)
            data["defects"].append({
                "id": get_next_id("D", data["defects"]),
                "model": sel_model_df,
                "date": datetime.now().strftime("%Y-%m-%d"),
                "time": datetime.now().strftime("%H:%M"),
                "process": df_process,
                "type": df_type,
                "count": df_count,
                "total": df_total,
                "rate": rate,
                "location": df_loc,
                "lot": df_lot,
                "description": df_desc,
                "status": "미조치",
                "person": df_person,
                "modified_by": df_person,
                "modified_at": datetime.now().strftime("%Y-%m-%d %H:%M")
            })
            st.success("✅ 불량 등록 완료!")
            st.rerun()

    model_df = [d for d in data["defects"] if d["model"] == sel_model_df]
    if model_df:
        st.divider()
        st.subheader(f"📊 {sel_model_df} 불량 이력 ({len(model_df)}건)")

        for i, d in enumerate(model_df):
            stat_icon = {"미조치":"🔴","분석중":"🟡","조치완료":"🟢","모니터링":"🔵"}.get(d["status"],"⚪")
            with st.expander(f"{stat_icon} [{d['id']}] {d['type']} | {d['process']} | {d['date']} | {d['count']}건 ({d.get('rate',0):.4f}%)"):
                dc1, dc2, dc3, dc4 = st.columns(4)
                with dc1:
                    st.write(f"**위치:** {d.get('location','-')}")
                with dc2:
                    st.write(f"**LOT:** {d.get('lot','-')}")
                with dc3:
                    st.write(f"**등록자:** {d.get('person','-')}")
                with dc4:
                    new_ds = st.selectbox("조치상태", ["미조치","분석중","조치완료","모니터링"],
                        index=["미조치","분석중","조치완료","모니터링"].index(d["status"]),
                        key=f"df_st_{d['id']}", label_visibility="collapsed")
                    if new_ds != d["status"]:
                        d["status"] = new_ds
                        d["modified_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")

                if d.get("description"):
                    st.write(f"📝 {d['description']}")
                st.caption(f"수정: {d.get('modified_at','')}")

        # 집계
        if len(model_df) >= 2:
            st.divider()
            ch1, ch2 = st.columns(2)
            with ch1:
                tc = {}
                for d in model_df:
                    tc[d["type"]] = tc.get(d["type"], 0) + d["count"]
                st.write("**불량유형별**")
                st.bar_chart(pd.DataFrame({"수량": tc}, index=tc.keys()))
            with ch2:
                pc = {}
                for d in model_df:
                    pc[d["process"]] = pc.get(d["process"], 0) + d["count"]
                st.write("**공정별**")
                st.bar_chart(pd.DataFrame({"수량": pc}, index=pc.keys()))
    else:
        st.info("등록된 불량이 없습니다.")


# ========== 탭4: BOM 변경 ==========
with tab4:
    st.header("📦 BOM 변경 이력")

    sel_model_bom = st.selectbox("모델 선택", model_names, key="bom_model")

    with st.form("add_bom", clear_on_submit=True):
        bom_c1, bom_c2 = st.columns(2)
        with bom_c1:
            bom_part = st.text_input("부품명/번호", placeholder="예: C105 MLCC 100nF 0402")
            bom_before = st.text_input("변경 전", placeholder="예: Samsung CL05A104KA5NNNC")
        with bom_c2:
            bom_reason = st.selectbox("변경사유", [
                "품질 개선", "원가 절감", "납기 이슈", "EOL 대응",
                "설계 변경", "고객 요청", "신뢰성 이슈", "기타"
            ])
            bom_after = st.text_input("변경 후", placeholder="예: Murata GRM155R61A104KE19D")

        bom_c3, bom_c4 = st.columns(2)
        with bom_c3:
            bom_approver = st.text_input("승인자", placeholder="이름", key="bom_approver")
        with bom_c4:
            bom_impact = st.text_input("영향분석", placeholder="변경으로 인한 영향 설명")

        if st.form_submit_button("➕ BOM 변경 등록", use_container_width=True):
            if bom_part:
                rev_num = len([b for b in data["bom_changes"] if b["model"] == sel_model_bom]) + 1
                data["bom_changes"].append({
                    "id": get_next_id("BOM", data["bom_changes"]),
                    "model": sel_model_bom,
                    "date": datetime.now().strftime("%Y-%m-%d"),
                    "rev": f"Rev.{rev_num}",
                    "part": bom_part,
                    "before": bom_before,
                    "after": bom_after,
                    "reason": bom_reason,
                    "approver": bom_approver,
                    "impact": bom_impact,
                    "modified_by": bom_approver,
                    "modified_at": datetime.now().strftime("%Y-%m-%d %H:%M")
                })
                st.success("✅ BOM 변경 등록!")
                st.rerun()

    model_bom = [b for b in data["bom_changes"] if b["model"] == sel_model_bom]
    if model_bom:
        st.divider()
        st.subheader(f"📋 {sel_model_bom} BOM 변경 이력 ({len(model_bom)}건)")
        df_bom = pd.DataFrame(model_bom)
        display_cols = ["id","date","rev","part","before","after","reason","approver","impact"]
        available = [c for c in display_cols if c in df_bom.columns]
        st.dataframe(df_bom[available], use_container_width=True, hide_index=True)
    else:
        st.info("등록된 BOM 변경이 없습니다.")


# ========== 탭5: FA 기록 ==========
with tab5:
    st.header("🔬 FA 기록")

    sel_model_fa = st.selectbox("모델 선택", model_names, key="fa_model")

    # 관련 불량 목록
    related_defects = [f"[{d['id']}] {d['type']} - {d.get('location','')}" for d in data["defects"] if d["model"] == sel_model_fa]

    with st.form("add_fa", clear_on_submit=True):
        fa_c1, fa_c2 = st.columns(2)
        with fa_c1:
            fa_title = st.text_input("FA 제목", placeholder="예: U23 BGA 솔더 크랙 분석")
            fa_related = st.selectbox("관련 불량", ["해당없음"] + related_defects) if related_defects else st.text_input("관련 불량ID", placeholder="예: D001")
        with fa_c2:
            fa_method = st.multiselect("분석방법", [
                "외관검사", "X-ray", "Cross Section", "SEM-EDS",
                "FT-IR", "DSC/TGA", "Pull/Shear Test", "C-SAM", "기타"
            ])
            fa_status = st.selectbox("진행상태", ["분석중", "완료", "보류"])

        fa_symptom = st.text_input("불량증상", placeholder="예: BGA 솔더 조인트 오픈")
        fa_cause = st.text_input("원인 (분석결과)", placeholder="예: 리플로우 피크온도 부족")
        fa_action = st.text_input("대책", placeholder="예: 피크온도 245→250°C 상향")
        fa_person = st.text_input("담당자", placeholder="이름", key="fa_person")

        if st.form_submit_button("➕ FA 기록 등록", use_container_width=True, type="primary"):
            if fa_title:
                related_id = ""
                if isinstance(fa_related, str) and fa_related.startswith("["):
                    related_id = fa_related.split("]")[0].replace("[","")
                data["fa_records"].append({
                    "id": get_next_id("FA", data["fa_records"]),
                    "model": sel_model_fa,
                    "related_defect": related_id,
                    "date": datetime.now().strftime("%Y-%m-%d"),
                    "title": fa_title,
                    "method": ", ".join(fa_method) if fa_method else "-",
                    "symptom": fa_symptom,
                    "cause": fa_cause,
                    "action": fa_action,
                    "status": fa_status,
                    "person": fa_person,
                    "modified_by": fa_person,
                    "modified_at": datetime.now().strftime("%Y-%m-%d %H:%M")
                })
                st.success("✅ FA 등록!")
                st.rerun()

    model_fa = [f for f in data["fa_records"] if f["model"] == sel_model_fa]
    if model_fa:
        st.divider()
        for fa in model_fa:
            s_icon = {"분석중":"🟡","완료":"🟢","보류":"🔴"}.get(fa["status"],"⚪")
            with st.expander(f"{s_icon} [{fa['id']}] {fa['title']}  |  {fa['date']}  |  {fa['status']}"):
                fc1, fc2 = st.columns(2)
                with fc1:
                    st.write(f"**분석방법:** {fa['method']}")
                    st.write(f"**불량증상:** {fa.get('symptom','-')}")
                    st.write(f"**담당자:** {fa.get('person','-')}")
                with fc2:
                    st.write(f"**원인:** {fa.get('cause','-') or '분석중...'}")
                    st.write(f"**대책:** {fa.get('action','-') or '수립중...'}")

                new_fas = st.selectbox("상태변경", ["분석중","완료","보류"],
                    index=["분석중","완료","보류"].index(fa["status"]),
                    key=f"fa_st_{fa['id']}")
                if new_fas != fa["status"]:
                    fa["status"] = new_fas
                    fa["modified_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                    st.rerun()

                st.caption(f"수정: {fa.get('modified_at','')} by {fa.get('modified_by','')}")
    else:
        st.info("등록된 FA 기록이 없습니다.")


# ========== 탭6: 모델 관리 ==========
with tab6:
    st.header("⚙️ 모델 관리")
    st.caption("모델 정보 수정, 단계 변경, 삭제 등")

    for m in data["models"]:
        with st.expander(f"📱 {m['name']}  ({m['id']})", expanded=False):
            mg1, mg2, mg3, mg4 = st.columns(4)
            with mg1:
                new_phase = st.selectbox("개발단계", ["Proto","EVT","DVT","PVT","MP"],
                    index=["Proto","EVT","DVT","PVT","MP"].index(m["phase"]),
                    key=f"mg_phase_{m['id']}")
                if new_phase != m["phase"]:
                    m["phase"] = new_phase
                    m["modified_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            with mg2:
                new_status = st.selectbox("상태", ["진행중","지연","이슈","완료","보류"],
                    index=["진행중","지연","이슈","완료","보류"].index(m["status"]),
                    key=f"mg_status_{m['id']}")
                if new_status != m["status"]:
                    m["status"] = new_status
                    m["modified_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            with mg3:
                new_issue = st.text_input("주요이슈", value=m.get("issue",""), key=f"mg_issue_{m['id']}")
                if new_issue != m.get("issue",""):
                    m["issue"] = new_issue
                    m["modified_at"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            with mg4:
                st.caption(f"생성: {m.get('created','')}")
                st.caption(f"수정: {m.get('modified_at','')}")

            if st.button(f"🗑️ {m['name']} 삭제", key=f"del_model_{m['id']}"):
                data["models"].remove(m)
                st.rerun()

# 하단
st.divider()
st.caption("SMT 개발 모델 PL 관리 대시보드 v2.0 | Streamlit Cloud Edition")
