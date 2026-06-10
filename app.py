import streamlit as st
import pandas as pd
import networkx as nx
import random
import io
import json
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Camp Group Creator", layout="wide")
st.sidebar.title("🏕️ Camp Group Creator")

# =========================================================================================
# ========================= GLOBAL DATA PROCESSING (SHARED BY ALL PAGES) ==================
# =========================================================================================
st.sidebar.header("📂 1. Core Data Upload")
responses_file = st.sidebar.file_uploader("Preference Survey (CSV)", type=["csv"])
students_file  = st.sidebar.file_uploader("Student List (CSV)",      type=["csv"])

# ── Tool selection comes FIRST so Y8 processing can be gated by page ─────────
st.sidebar.markdown("---")
st.sidebar.header("🎯 2. Select Tool")
page = st.sidebar.radio("", ["🏕️ Y8 Group Creator", "🏔️ Y9 Journey Groups", "📋 Final Roster & Leader Builder"])

# ── Y8 global data processing (only when a Y8 tool is active) ────────────────
df_merged_full = None
df_stud_pub = None
mtb_ability_col = None
all_students_list = []

if responses_file and students_file and page in ["🏕️ Y8 Group Creator", "📋 Final Roster & Leader Builder"]:
    responses_file.seek(0)
    students_file.seek(0)
    df_resp = pd.read_csv(responses_file)
    df_stud = pd.read_csv(students_file)

    # Clean data
    df_resp['Email address'] = df_resp['Email address'].astype(str).str.strip().str.lower()
    df_stud['Email'] = df_stud['Email'].astype(str).str.strip().str.lower()
    df_stud['Official Name'] = df_stud['Preferred name'].astype(str).str.strip() + " " + df_stud['Surname'].astype(str).str.strip()

    if 'Gender' in df_stud.columns: df_stud['Gender'] = df_stud['Gender'].astype(str).str.strip().str.lower()
    else: df_stud['Gender'] = 'o' 

    # Explicitly map the "Code" column to "Student ID"
    if 'Code' in df_stud.columns:
        df_stud['Student ID'] = df_stud['Code'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    else:
        id_col = next((c for c in df_stud.columns if 'id' in str(c).lower() and 'email' not in str(c).lower()), None)
        if id_col: df_stud['Student ID'] = df_stud[id_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        else: df_stud['Student ID'] = "N/A"

    df_stud_pub = df_stud.copy()

    df_merged_full = pd.merge(df_stud[['Email', 'Official Name', 'Rollgroup', 'Gender', 'Student ID']], 
                         df_resp, left_on='Email', right_on='Email address', how='left')

    df_merged_full['Responded'] = df_merged_full['Email address'].notna()

    valid_classes = df_resp['Which Connections class are you in? '].dropna().unique()
    rg_to_class = {f"8{str(c).strip()[0].upper()}": str(c).strip() for c in valid_classes}
    df_merged_full['Which Connections class are you in? '] = df_merged_full['Which Connections class are you in? '].fillna(
        df_merged_full['Rollgroup'].astype(str).str.strip().str.upper().map(rg_to_class)
    )
    df_merged_full = df_merged_full[df_merged_full['Which Connections class are you in? '].isin(valid_classes)]

    mtb_ability_col = next((c for c in df_merged_full.columns if 'comfortable are you riding a bike' in str(c).lower()), None)
    all_students_list = sorted(df_merged_full['Official Name'].dropna().unique().tolist())


# =========================================================================================
# ========================= PAGE 1: Y8 GROUP CREATOR ======================================
# =========================================================================================
if page == "🏕️ Y8 Group Creator":
    st.title("🏕️ Year 8 Camp Group Creator")

    # Default AI scoring weights
    default_weights = {
        "penalty_group_size_diff_multiplier": 10, "penalty_group_over_14": 500000, "penalty_forced_location": 1000000,
        "penalty_must_go_with": 1000000, "penalty_sole_gender": 200000, "penalty_zero_friends": 100000,
        "penalty_minority_gender_no_friends": 100000, "penalty_separation": 100000, "penalty_veto_activity": 50000,
        "penalty_minority_gender_with_friends": 1000, "reward_friend_1": 120, "reward_friend_2": 100, 
        "reward_activity_score_multiplier": 100
    }

    # Initialize session states
    if 'separation_pairs' not in st.session_state: st.session_state.separation_pairs = []
    if 'must_go_pairs' not in st.session_state: st.session_state.must_go_pairs = []
    if 'forced_locations' not in st.session_state: st.session_state.forced_locations = {}
    if 'not_attending' not in st.session_state: st.session_state.not_attending = []
    if 'generation_seed' not in st.session_state: st.session_state.generation_seed = 0
    if 'weights' not in st.session_state: st.session_state.weights = default_weights.copy()
    if 'optimized_groups' not in st.session_state: st.session_state.optimized_groups = {}
    if 'class_opt_states' not in st.session_state: st.session_state.class_opt_states = {}

    if df_merged_full is not None:
        
        # --- LOAD SETTINGS FEATURE ---
        st.sidebar.markdown("---")
        st.sidebar.header("📂 Load Rules & Weights")
        
        settings_file = st.sidebar.file_uploader("Load Saved JSON", type=["json"])
        
        if settings_file is not None:
            if st.session_state.get('last_loaded_file') != settings_file.file_id:
                try:
                    settings = json.load(settings_file)
                    st.session_state.separation_pairs = [tuple(x) for x in settings.get('separation_pairs', [])]
                    st.session_state.must_go_pairs = [tuple(x) for x in settings.get('must_go_pairs', [])]
                    st.session_state.forced_locations = settings.get('forced_locations', {})
                    
                    loaded_na = [s for s in settings.get('not_attending', []) if s in all_students_list]
                    st.session_state.not_attending = loaded_na
                    
                    if 'weights' in settings:
                        st.session_state.weights.update(settings['weights'])
                        if 'reward_friend_1' not in st.session_state.weights: st.session_state.weights['reward_friend_1'] = 120
                        if 'reward_friend_2' not in st.session_state.weights: st.session_state.weights['reward_friend_2'] = 100
                    
                    st.session_state.last_loaded_file = settings_file.file_id
                    st.sidebar.success("Settings loaded successfully!")
                    st.rerun() 
                except Exception as e:
                    st.sidebar.error("Error loading the rules file.")

        st.sidebar.markdown("---")

        # --- ADVANCED AI WEIGHTS PANEL ---
        with st.sidebar.expander("⚖️ Advanced AI Weights"):
            st.write("Adjust how strongly the AI penalizes bad groups or rewards good ones.")
            w = st.session_state.weights
            w['penalty_forced_location'] = st.number_input("Forced Location Broken (Penalty)", value=w['penalty_forced_location'], step=10000)
            w['penalty_must_go_with'] = st.number_input("'Must Go With' Broken (Penalty)", value=w['penalty_must_go_with'], step=10000)
            w['penalty_sole_gender'] = st.number_input("Sole Gender in Group (Penalty)", value=w['penalty_sole_gender'], step=10000)
            w['penalty_zero_friends'] = st.number_input("Zero Friends in Group (Penalty)", value=w['penalty_zero_friends'], step=10000)
            w['penalty_minority_gender_no_friends'] = st.number_input("Minority Gender + No Friends (Penalty)", value=w['penalty_minority_gender_no_friends'], step=10000)
            w['penalty_separation'] = st.number_input("'Can't Go With' Broken (Penalty)", value=w['penalty_separation'], step=10000)
            w['penalty_veto_activity'] = st.number_input("Activity Veto '1' Ignored (Penalty)", value=w['penalty_veto_activity'], step=10000)
            w['penalty_minority_gender_with_friends'] = st.number_input("Minority Gender + Has Friends (Penalty)", value=w['penalty_minority_gender_with_friends'], step=100)
            w['reward_activity_score_multiplier'] = st.number_input("Activity Score Multiplier (Preference Weight)", value=w['reward_activity_score_multiplier'], step=10)
            
            w['reward_friend_1'] = st.number_input("Reward for Keeping Friend 1", value=w.get('reward_friend_1', 120), step=10)
            w['reward_friend_2'] = st.number_input("Reward for Keeping Friend 2+", value=w.get('reward_friend_2', 100), step=10)
            
            w['penalty_group_size_diff_multiplier'] = st.number_input("Group Size Imbalance Multiplier", value=w['penalty_group_size_diff_multiplier'], step=1)
            
            if st.button("🔄 Reset Weights to Defaults"):
                st.session_state.weights = default_weights.copy()
                st.rerun()

        st.sidebar.markdown("---")

        # --- UI: NOT ATTENDING ---
        st.sidebar.header("🚫 Not Attending")
        not_attending_list = st.sidebar.multiselect("Select students:", options=all_students_list, default=st.session_state.not_attending)
        st.session_state.not_attending = not_attending_list

        df_merged = df_merged_full[~df_merged_full['Official Name'].isin(not_attending_list)].copy()
        
        missing_students_df = df_stud[~df_stud['Email'].isin(df_resp['Email address'])].copy()
        missing_students_df = missing_students_df[~missing_students_df['Official Name'].isin(not_attending_list)]

        attending_students_list = sorted(df_merged['Official Name'].dropna().unique().tolist())

        # --- FIND STUDENT & REGENERATE BAR ---
        st.markdown("---")
        col_search, col_regen = st.columns([2, 1])
        with col_search:
            st.subheader("🔍 Find a Student")
            student_to_find = st.selectbox("Search for a student to quickly see their placement:", [""] + attending_students_list, key="find_student")
        with col_regen:
            st.subheader("⚙️ AI Controls")
            search_depth = st.select_slider("Search Depth (Time vs Perfection):", options=["Fast (1-2s)", "Standard (5-10s)", "Deep Search (20-40s)"], value="Standard (5-10s)")
            if st.button("🔄 Regenerate AI Groups", help="Forces the AI to calculate a completely new arrangement", use_container_width=True):
                st.session_state.generation_seed += 1
                st.rerun()

        find_result_container = st.empty()
        st.markdown("---")

        # --- UI: FORCE LOCATION ---
        st.sidebar.header("📍 Force Camp Location")
        force_stud = st.sidebar.selectbox("Student to Force", [""] + attending_students_list, key="f_stud")
        force_camp = st.sidebar.selectbox("Select Camp", ["", "Freycinet", "Bay of Fires"], key="f_camp")
        if st.sidebar.button("Add Location Rule"):
            if force_stud and force_camp:
                st.session_state.forced_locations[force_stud] = force_camp
                st.rerun() 
                
        if st.session_state.forced_locations:
            st.sidebar.write("**Active Location Rules:**")
            for s, c in list(st.session_state.forced_locations.items()): 
                col1, col2 = st.sidebar.columns([5, 1])
                col1.write(f"- {s} ➡️ {c}")
                if col2.button("❌", key=f"del_loc_{s}", help="Delete rule"):
                    del st.session_state.forced_locations[s]
                    st.rerun()
            if st.sidebar.button("Clear Location Rules"): st.session_state.forced_locations = {}; st.rerun()

        # --- UI: MUST GO WITH ---
        st.sidebar.header("🤝 'Must Go With' Rules")
        col3, col4 = st.sidebar.columns(2)
        with col3: must_1 = st.selectbox("Student A", [""] + attending_students_list, key="m1")
        with col4: must_2 = st.selectbox("Student B", [""] + attending_students_list, key="m2")
        if st.sidebar.button("Add 'Must Go' Rule"):
            if must_1 and must_2 and must_1 != must_2:
                st.session_state.must_go_pairs.append((must_1, must_2))
                st.rerun() 
                
        if st.session_state.must_go_pairs:
            st.sidebar.write("**Active 'Must Go' Rules:**")
            for i, pair in enumerate(st.session_state.must_go_pairs): 
                col1, col2 = st.sidebar.columns([5, 1])
                col1.write(f"- {pair[0]} & {pair[1]}")
                if col2.button("❌", key=f"del_mg_{i}", help="Delete rule"):
                    st.session_state.must_go_pairs.pop(i)
                    st.rerun()
            if st.sidebar.button("Clear 'Must Go' Rules"): st.session_state.must_go_pairs = []; st.rerun()

        # --- UI: CAN'T GO WITH ---
        st.sidebar.header("⚠️ 'Can't Go With' Rules")
        col5, col6 = st.sidebar.columns(2)
        with col5: sep_1 = st.selectbox("Student 1", [""] + attending_students_list, key="s1")
        with col6: sep_2 = st.selectbox("Student 2", [""] + attending_students_list, key="s2")
        if st.sidebar.button("Add 'Can't Go' Rule"):
            if sep_1 and sep_2 and sep_1 != sep_2:
                st.session_state.separation_pairs.append((sep_1, sep_2))
                st.rerun() 
                
        if st.session_state.separation_pairs:
            st.sidebar.write("**Active 'Can't Go' Rules:**")
            for i, pair in enumerate(st.session_state.separation_pairs): 
                col1, col2 = st.sidebar.columns([5, 1])
                col1.write(f"- {pair[0]} ≠ {pair[1]}")
                if col2.button("❌", key=f"del_sep_{i}", help="Delete rule"):
                    st.session_state.separation_pairs.pop(i)
                    st.rerun()
            if st.sidebar.button("Clear 'Can't Go' Rules"): st.session_state.separation_pairs = []; st.rerun()

        # --- SAVE SETTINGS FEATURE ---
        st.sidebar.markdown("---")
        st.sidebar.header("💾 Save Your Setup")
        
        current_settings = {
            "separation_pairs": st.session_state.separation_pairs,
            "must_go_pairs": st.session_state.must_go_pairs,
            "forced_locations": st.session_state.forced_locations,
            "not_attending": st.session_state.not_attending,
            "weights": st.session_state.weights 
        }
        st.sidebar.download_button(
            label="📥 Save Rules & Weights for Next Time (JSON)",
            data=json.dumps(current_settings, indent=4),
            file_name="camp_rules.json",
            mime="application/json"
        )

        # --- 3. CALCULATE ACTIVITY SCORES ---
        df_merged['Freycinet Score'] = pd.to_numeric(df_merged.get('How excited would you be to go sea kayaking?', 0), errors='coerce').fillna(0) + \
                                       pd.to_numeric(df_merged.get('How excited would you be to go coasteering?', 0), errors='coerce').fillna(0)
        df_merged['BoF Score'] = pd.to_numeric(df_merged.get('How excited would you be to go mountain bike riding on Blue (intermediate) flowy trails?', 0), errors='coerce').fillna(0) + \
                                 pd.to_numeric(df_merged.get('How excited would you be to go snorkelling?', 0), errors='coerce').fillna(0)
        
        def get_preference(row):
            if not row['Responded']: return 'No Response (Tie)'
            if row['Freycinet Score'] > row['BoF Score']: return 'Freycinet'
            elif row['BoF Score'] > row['Freycinet Score']: return 'Bay of Fires'
            else: return 'Tie'
        df_merged['Camp Preference'] = df_merged.apply(get_preference, axis=1)

        # --- 4. OPTIMIZATION ALGORITHM PER CLASS ---
        classes = sorted(df_merged['Which Connections class are you in? '].dropna().unique().tolist())
        styled_class_dfs = {}
        leader_overview_data = []
        isolated_students = []

        if not_attending_list:
            st.warning(f"🚫 {len(not_attending_list)} student(s) marked as Not Attending. They have been removed from the sorting process.")

        st.header("Camp Group Assignments")

        # --- CROSS-CLASS RULE WARNING ---
        student_to_class = df_merged.set_index('Official Name')['Which Connections class are you in? '].to_dict()
        cross_class_warnings = []
        for s1, s2 in st.session_state.must_go_pairs + st.session_state.separation_pairs:
            c1 = student_to_class.get(s1)
            c2 = student_to_class.get(s2)
            if c1 and c2 and c1 != c2:
                cross_class_warnings.append(f"{s1} ({c1}) & {s2} ({c2})")
                
        if cross_class_warnings:
            st.warning("⚠️ **Warning: Rules across different classes detected!** The optimizer sorts class-by-class and cannot move a student to a different class. The following rules span multiple classes and will be ignored by the algorithm: " + ", ".join(cross_class_warnings))

        for cls in classes:
            class_df = df_merged[df_merged['Which Connections class are you in? '] == cls].copy()
            class_students = set(class_df['Official Name'])
            
            class_sep = [p for p in st.session_state.separation_pairs if p[0] in class_students or p[1] in class_students]
            class_must = [p for p in st.session_state.must_go_pairs if p[0] in class_students or p[1] in class_students]
            class_force = {k: v for k, v in st.session_state.forced_locations.items() if k in class_students}
            class_na = [s for s in st.session_state.not_attending if s in class_students]
            
            current_class_state = str({
                "sep": class_sep,
                "must": class_must,
                "force": class_force,
                "na": class_na,
                "w": st.session_state.weights,
                "seed": st.session_state.generation_seed,
                "depth": search_depth,
                "data_len": len(class_df)
            })

            if st.session_state.class_opt_states.get(cls) != current_class_state:
                if cls in st.session_state.optimized_groups:
                    del st.session_state.optimized_groups[cls] 
                st.session_state.class_opt_states[cls] = current_class_state

            G = nx.DiGraph()
            friend_requests = {}
            friend_cols = [c for c in class_df.columns if 'classmate' in c.lower()]
            
            for idx, row in class_df.iterrows():
                student = row['Official Name']
                
                sk = pd.to_numeric(row.get('How excited would you be to go sea kayaking?', 0), errors='coerce')
                coast = pd.to_numeric(row.get('How excited would you be to go coasteering?', 0), errors='coerce')
                mtb = pd.to_numeric(row.get('How excited would you be to go mountain bike riding on Blue (intermediate) flowy trails?', 0), errors='coerce')
                snork = pd.to_numeric(row.get('How excited would you be to go snorkelling?', 0), errors='coerce')
                
                f_veto = True if (sk == 1 or coast == 1) else False
                b_veto = True if (mtb == 1 or snork == 1) else False
                
                G.add_node(student, pref=row['Camp Preference'], f_score=row['Freycinet Score'], b_score=row['BoF Score'], gender=row.get('Gender', 'o'), f_veto=f_veto, b_veto=b_veto)
                
                reqs = []
                if row['Responded']:
                    for col in friend_cols:
                        friend_name = row.get(col)
                        if pd.notna(friend_name) and isinstance(friend_name, str) and friend_name.strip() != "" and friend_name.strip() not in not_attending_list:
                            reqs.append(friend_name.strip())
                friend_requests[student] = reqs
                
            for s1, s2 in st.session_state.must_go_pairs:
                if s1 in class_students and s2 in class_students:
                    if s1 in friend_requests and s2 not in friend_requests[s1]: friend_requests[s1].append(s2)
                    if s2 in friend_requests and s1 not in friend_requests[s2]: friend_requests[s2].append(s1)
                    
            for s1, s2 in st.session_state.separation_pairs:
                if s1 in friend_requests and s2 in friend_requests[s1]: friend_requests[s1].remove(s2)
                if s2 in friend_requests and s1 in friend_requests[s2]: friend_requests[s2].remove(s1)

            for student, reqs in friend_requests.items():
                valid_reqs = [f for f in reqs if f in class_students]
                friend_requests[student] = valid_reqs 
                for f_name in valid_reqs: 
                    G.add_edge(student, f_name)

            students = sorted(list(G.nodes))

            if cls in st.session_state.optimized_groups:
                freycinet_group, bof_group = st.session_state.optimized_groups[cls]
            else:
                def calc_score(f_set, b_set):
                    w = st.session_state.weights 
                    score = 0
                    diff = abs(len(f_set) - len(b_set))
                    score -= (diff * diff * w['penalty_group_size_diff_multiplier']) 
                    if len(f_set) > 14 or len(b_set) > 14: score -= w['penalty_group_over_14']
                    
                    for f_stud, f_camp in st.session_state.forced_locations.items():
                        if f_camp == 'Freycinet' and f_stud in b_set: score -= w['penalty_forced_location']
                        elif f_camp == 'Bay of Fires' and f_stud in f_set: score -= w['penalty_forced_location']

                    for s1, s2 in st.session_state.must_go_pairs:
                        if (s1 in f_set and s2 in b_set) or (s1 in b_set and s2 in f_set): score -= w['penalty_must_go_with']

                    for s1, s2 in st.session_state.separation_pairs:
                        if (s1 in f_set and s2 in f_set) or (s1 in b_set and s2 in b_set): score -= w['penalty_separation']

                    for set_group in [f_set, b_set]:
                        m_count = sum(1 for s in set_group if G.nodes[s]['gender'] == 'm')
                        f_count = sum(1 for s in set_group if G.nodes[s]['gender'] == 'f')
                        for s in set_group:
                            is_f = set_group == f_set
                            
                            base_act_score = G.nodes[s]['f_score'] if is_f else G.nodes[s]['b_score']
                            score += (base_act_score * w['reward_activity_score_multiplier'])
                            
                            if is_f and G.nodes[s]['f_veto']: score -= w['penalty_veto_activity']
                            if not is_f and G.nodes[s]['b_veto']: score -= w['penalty_veto_activity']
                            
                            reqs = friend_requests.get(s, [])
                            friends_in_group = sum(1 for f in reqs if f in set_group) if reqs else 0
                            
                            if reqs:
                                if friends_in_group == 0: 
                                    score -= w['penalty_zero_friends'] 
                                else:
                                    for i, friend_name in enumerate(reqs):
                                        if friend_name in set_group:
                                            if i == 0: score += w.get('reward_friend_1', 120)
                                            else: score += w.get('reward_friend_2', 100)
                                
                            my_gen = G.nodes[s]['gender']
                            my_gen_count = m_count if my_gen == 'm' else (f_count if my_gen == 'f' else 0)
                            if my_gen in ['m', 'f']:
                                if my_gen_count == 1: score -= w['penalty_sole_gender'] 
                                elif 1 < my_gen_count < 4:
                                    if friends_in_group == 0: score -= w['penalty_minority_gender_no_friends'] 
                                    else: score -= w['penalty_minority_gender_with_friends'] 
                    return score

                best_overall_f = set()
                best_overall_b = set()
                best_overall_score = -float('inf')

                if "Fast" in search_depth: restarts, swaps = 50, 1000
                elif "Standard" in search_depth: restarts, swaps = 150, 2500
                else: restarts, swaps = 400, 3000

                with st.spinner(f"AI Optimizing Class {cls} ({restarts * swaps:,} checks)..."):
                    random.seed(f"camp_{cls}_fixed_{st.session_state.generation_seed}") 
                    
                    for restart in range(restarts):
                        shuffled_students = students.copy()
                        random.shuffle(shuffled_students)
                        curr_f = set(shuffled_students[:len(shuffled_students)//2])
                        curr_b = set(shuffled_students[len(shuffled_students)//2:])
                        curr_score = calc_score(curr_f, curr_b)
                        
                        for _ in range(swaps):
                            new_f, new_b = set(curr_f), set(curr_b)
                            if random.random() < 0.7 and len(new_f) > 0 and len(new_b) > 0:
                                s1 = random.choice(sorted(list(new_f)))
                                s2 = random.choice(sorted(list(new_b)))
                                new_f.remove(s1); new_f.add(s2)
                                new_b.remove(s2); new_b.add(s1)
                            else:
                                if random.random() < 0.5 and len(new_f) > 0:
                                    s = random.choice(sorted(list(new_f)))
                                    new_f.remove(s); new_b.add(s)
                                elif len(new_b) > 0:
                                    s = random.choice(sorted(list(new_b)))
                                    new_b.remove(s); new_f.add(s)
                            
                            new_score = calc_score(new_f, new_b)
                            if new_score > curr_score:
                                curr_score = new_score
                                curr_f, curr_b = new_f, new_b
                                
                        if curr_score > best_overall_score:
                            best_overall_score = curr_score
                            best_overall_f = curr_f
                            best_overall_b = curr_b

                freycinet_group = sorted(list(best_overall_f))
                bof_group = sorted(list(best_overall_b))
                st.session_state.optimized_groups[cls] = (freycinet_group, bof_group)

            # --- FIND STUDENT RESULT ---
            if student_to_find in freycinet_group:
                find_result_container.success(f"🎯 **{student_to_find}** was placed in **Freycinet** (Class {cls}). Scroll down to see their group!")
            elif student_to_find in bof_group:
                find_result_container.success(f"🎯 **{student_to_find}** was placed in **Bay of Fires** (Class {cls}). Scroll down to see their group!")

            # --- PREPARE DATAFRAME WITH VISUAL SPLITS ---
            class_roster = []
            
            for camp_name, camp_group in [('Freycinet', freycinet_group), ('Bay of Fires', bof_group)]:
                if camp_name == 'Bay of Fires':
                    class_roster.append({
                        'Student ID': '---', 'Email': '---', 'Student': f'--- BAY OF FIRES GROUP ---', 'Responded': '---', 'Gender': '---', 'Friend 1': '---', 'Friend 2': '---',
                        'Sea Kayak': '---', 'Coasteer': '---', 'Freycinet TOTAL': '---',
                        'MTB': '---', 'Snorkel': '---', 'Bay of Fires TOTAL': '---',
                        'Preferred Camp': '---', 'Assigned Camp': '---'
                    })
                    leader_overview_data.append({
                        'Class': f"--- {cls} ---", 'Assigned Camp': '--- BAY OF FIRES ---', 'Student ID': '---', 'Email': '---', 'Student': '---------------------------',
                        'Responded': '---', 'Sea Kayak': '---', 'Coasteer': '---', 'MTB Interest': '---', 'Snorkel': '---',
                        'General Camping Skill': '---', 'Sleeping Outdoors/Bugs': '---', 'Swimming Confidence': '---', 'Bike Comfort': '---', 'Overnight Hike': '---',
                        'Reaction to Hardship': '---', 'Gear Independence': '---', 'Group Teamwork': '---'
                    })
                else:
                    leader_overview_data.append({
                        'Class': f"--- {cls} ---", 'Assigned Camp': '--- FREYCINET ---', 'Student ID': '---', 'Email': '---', 'Student': '---------------------------',
                        'Responded': '---', 'Sea Kayak': '---', 'Coasteer': '---', 'MTB Interest': '---', 'Snorkel': '---',
                        'General Camping Skill': '---', 'Sleeping Outdoors/Bugs': '---', 'Swimming Confidence': '---', 'Bike Comfort': '---', 'Overnight Hike': '---',
                        'Reaction to Hardship': '---', 'Gear Independence': '---', 'Group Teamwork': '---'
                    })

                for student in camp_group:
                    student_row = class_df[class_df['Official Name'] == student].iloc[0]
                    reqs = friend_requests.get(student, [])
                    friend1 = reqs[0] if len(reqs) > 0 else ""
                    friend2 = reqs[1] if len(reqs) > 1 else ""
                    
                    friends_in_camp = sum(1 for f in reqs if f in camp_group)
                    if reqs and friends_in_camp == 0:
                        isolated_students.append({'Class': cls, 'Camp': camp_name, 'Student': student, 'Reason': 'Got 0 requested friends'})
                    elif not reqs:
                        isolated_students.append({'Class': cls, 'Camp': camp_name, 'Student': student, 'Reason': 'Made no friend requests'})

                    sk = student_row.get('How excited would you be to go sea kayaking?', 'N/A') if student_row['Responded'] else 'N/A'
                    coast = student_row.get('How excited would you be to go coasteering?', 'N/A') if student_row['Responded'] else 'N/A'
                    mtb_int = student_row.get('How excited would you be to go mountain bike riding on Blue (intermediate) flowy trails?', 'N/A') if student_row['Responded'] else 'N/A'
                    snork = student_row.get('How excited would you be to go snorkelling?', 'N/A') if student_row['Responded'] else 'N/A'
                    mtb_ab = student_row.get(mtb_ability_col, 'N/A') if mtb_ability_col and student_row['Responded'] else 'N/A'
                    
                    class_roster.append({
                        'Student ID': student_row.get('Student ID', 'N/A'),
                        'Email': student_row.get('Email', 'N/A'),
                        'Student': student,
                        'Responded': 'Yes' if student_row['Responded'] else 'No',
                        'Gender': G.nodes[student]['gender'].upper(),
                        'Friend 1': friend1,
                        'Friend 2': friend2,
                        'Sea Kayak': sk,
                        'Coasteer': coast,
                        'Freycinet TOTAL': G.nodes[student]['f_score'],
                        'MTB': mtb_int,
                        'Snorkel': snork,
                        'Bay of Fires TOTAL': G.nodes[student]['b_score'],
                        'Preferred Camp': G.nodes[student]['pref'],
                        'Assigned Camp': camp_name
                    })
                    
                    leader_overview_data.append({
                        'Class': cls, 'Assigned Camp': camp_name, 
                        'Student ID': student_row.get('Student ID', 'N/A'),
                        'Email': student_row.get('Email', 'N/A'),
                        'Student': student,
                        'Responded': 'Yes' if student_row['Responded'] else 'No',
                        
                        'Sea Kayak': sk if camp_name == 'Freycinet' else '---',
                        'Coasteer': coast if camp_name == 'Freycinet' else '---',
                        'MTB Interest': mtb_int if camp_name == 'Bay of Fires' else '---',
                        'Snorkel': snork if camp_name == 'Bay of Fires' else '---',
                        
                        'General Camping Skill': student_row.get('How confident are you in your general camping skills (like setting up a tent or cooking on a camp stove)?', 'N/A') if student_row['Responded'] else 'N/A',
                        'Sleeping Outdoors/Bugs': student_row.get('How do you feel about sleeping outdoors, using long-drop toilets, and being around dirt and bugs for a few days?', 'N/A') if student_row['Responded'] else 'N/A',
                        'Swimming Confidence': student_row.get('How confident are you swimming in deep water where you cannot touch the bottom?', 'N/A') if student_row['Responded'] else 'N/A',
                        'Bike Comfort': mtb_ab,
                        'Overnight Hike': student_row.get('How excited are you to go on an overnight hike carrying your gear?', 'N/A') if student_row['Responded'] else 'N/A',
                        'Reaction to Hardship': student_row.get('When a physical activity gets tiring, difficult, or the weather turns bad, how do you usually react?', 'N/A') if student_row['Responded'] else 'N/A',
                        'Gear Independence': student_row.get('How good are you at organising your own gear, packing your bag, and keeping yourself comfortable and safe without teachers or parents reminding you?', 'N/A') if student_row['Responded'] else 'N/A',
                        'Group Teamwork': student_row.get('How confident are you in helping your group work together, joining in discussions, and making classmates feel included?', 'N/A') if student_row['Responded'] else 'N/A'
                    })
            
            leader_overview_data.append({
                'Class': "", 'Assigned Camp': "", 'Student ID': "", 'Email': "", 'Student': "", 'Responded': "", 
                'Sea Kayak': "", 'Coasteer': "", 'MTB Interest': "", 'Snorkel': "",
                'General Camping Skill': "", 'Sleeping Outdoors/Bugs': "", 'Swimming Confidence': "", 'Bike Comfort': "", 'Overnight Hike': "",
                'Reaction to Hardship': "", 'Gear Independence': "", 'Group Teamwork': ""
            })
                
            df_class_final = pd.DataFrame(class_roster)

            # --- HIGHLIGHTING LOGIC ---
            def highlight_cells(row, df_reference, fr_dict):
                colors = [''] * len(row)
                
                if row['Student'] == student_to_find:
                    return ['background-color: #85e085; font-weight: bold; color: black; border: 2px solid green'] * len(row)

                if '---' in str(row['Student']):
                    return ['background-color: #d9d9d9; font-weight: bold; color: black; text-align: center'] * len(row)
                    
                student_to_camp = df_reference.set_index('Student')['Assigned Camp'].to_dict()
                
                if row['Responded'] == 'No': colors[df_reference.columns.get_loc('Responded')] = 'background-color: #ffffcc'
                
                for act_col in ['Sea Kayak', 'Coasteer', 'MTB', 'Snorkel']:
                    col_idx = df_reference.columns.get_loc(act_col)
                    val = row.get(act_col)
                    try:
                        if float(val) == 1.0:
                            if act_col in ['Sea Kayak', 'Coasteer'] and row['Assigned Camp'] == 'Freycinet':
                                colors[col_idx] = 'background-color: #ff4d4d; color: white; font-weight: bold'
                            elif act_col in ['MTB', 'Snorkel'] and row['Assigned Camp'] == 'Bay of Fires':
                                colors[col_idx] = 'background-color: #ff4d4d; color: white; font-weight: bold'
                    except:
                        pass

                camp_idx = df_reference.columns.get_loc('Assigned Camp')
                if row['Preferred Camp'] not in ['Tie', 'No Response (Tie)'] and row['Assigned Camp'] != row['Preferred Camp']:
                    colors[camp_idx] = 'background-color: #ffcccc' 
                    
                for f_col in ['Friend 1', 'Friend 2']:
                    f_idx = df_reference.columns.get_loc(f_col)
                    if row[f_col] != "":
                        f_camp = student_to_camp.get(row[f_col])
                        if f_camp is not None and f_camp != row['Assigned Camp']: colors[f_idx] = 'background-color: #ffcccc'
                            
                student_idx = df_reference.columns.get_loc('Student')
                
                reqs = fr_dict.get(row['Student'], [])
                if reqs:
                    friends_in_camp = sum(1 for f in reqs if student_to_camp.get(f) == row['Assigned Camp'])
                    if friends_in_camp == 0: colors[student_idx] = 'background-color: #ff9900; color: black; font-weight: bold'
                
                for s1, s2 in st.session_state.separation_pairs:
                    if row['Student'] == s1 and student_to_camp.get(s2) == row['Assigned Camp']: colors[student_idx] = 'background-color: #ff4d4d'
                    elif row['Student'] == s2 and student_to_camp.get(s1) == row['Assigned Camp']: colors[student_idx] = 'background-color: #ff4d4d'

                my_gender = row['Gender']
                if my_gender in ['M', 'F']:
                    gender_count_in_camp = sum((df_reference['Assigned Camp'] == row['Assigned Camp']) & (df_reference['Gender'] == my_gender))
                    gen_idx = df_reference.columns.get_loc('Gender')
                    if gender_count_in_camp == 1:
                        colors[gen_idx] = 'background-color: #ff4d4d; color: white' 
                    elif 1 < gender_count_in_camp < 4:
                        colors[gen_idx] = 'background-color: #ffcc00; color: black' 

                return colors

            styled_df = df_class_final.style.apply(highlight_cells, df_reference=df_class_final, fr_dict=friend_requests, axis=1)
            styled_class_dfs[cls] = styled_df

            st.subheader(f"Class: {cls}")
            f_males = sum(1 for s in freycinet_group if G.nodes[s]['gender'] == 'm')
            f_females = sum(1 for s in freycinet_group if G.nodes[s]['gender'] == 'f')
            b_males = sum(1 for s in bof_group if G.nodes[s]['gender'] == 'm')
            b_females = sum(1 for s in bof_group if G.nodes[s]['gender'] == 'f')
            
            st.write(f"**Freycinet:** {len(freycinet_group)} students ({f_males}M, {f_females}F) | **Bay of Fires:** {len(bof_group)} students ({b_males}M, {b_females}F)")
            st.dataframe(styled_df, use_container_width=True)

        # --- 5. DASHBOARDS ---
        colA, colB = st.columns(2)
        with colA:
            st.write("### ⚠️ At Risk of Isolation")
            if isolated_students: st.dataframe(pd.DataFrame(isolated_students), use_container_width=True)
            else: st.success("No isolated students!")
                
        with colB:
            st.write("### ❌ Missing Responses List")
            if not missing_students_df.empty: st.dataframe(missing_students_df[['First name', 'Surname', 'Rollgroup', 'Email']], use_container_width=True)
            else: st.success("Everyone has responded!")

        # --- LEADER OVERVIEW STYLING ---
        df_leaders = pd.DataFrame(leader_overview_data)
        
        def highlight_leader_overview(row, df_ref):
            colors = [''] * len(row)
            
            if '---' in str(row.get('Student', '')):
                return ['background-color: #d9d9d9; font-weight: bold; text-align: center'] * len(row)
            
            cols_to_check = [
                'General Camping Skill', 'Sleeping Outdoors/Bugs', 'Swimming Confidence', 'Bike Comfort', 'Overnight Hike', 
                'Reaction to Hardship', 'Gear Independence', 'Group Teamwork',
                'Sea Kayak', 'Coasteer', 'MTB Interest', 'Snorkel'
            ]

            for col_name in cols_to_check:
                if col_name in df_ref.columns:
                    col_idx = df_ref.columns.get_loc(col_name)
                    val = str(row[col_name]).strip()
                    match = re.search(r'^\s*([1-5])(?:\D|$)', val)
                    if match:
                        num = int(match.group(1))
                        if num in [1, 2]: colors[col_idx] = 'background-color: #a83232; color: white; font-weight: bold' # Dark Red
                        elif num in [3, 4]: colors[col_idx] = 'background-color: #ff9933; color: black; font-weight: bold' # Orange
                        elif num == 5: colors[col_idx] = 'background-color: #ffcc00; color: black; font-weight: bold' # Yellow
            return colors

        styled_leaders = df_leaders.style.apply(highlight_leader_overview, df_ref=df_leaders, axis=1)

        # --- 6. EXCEL EXPORT ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for cls, styled_df in styled_class_dfs.items():
                styled_df.to_excel(writer, sheet_name=str(cls)[:31], index=False)
                worksheet = writer.sheets[str(cls)[:31]]
                for col in worksheet.columns: worksheet.column_dimensions[col[0].column_letter].width = max((len(str(cell.value)) for cell in col), default=0) + 2

            # LEADER OVERVIEW
            styled_leaders.to_excel(writer, sheet_name="Leader Overview", index=False)
            worksheet_l = writer.sheets["Leader Overview"]
            for col in worksheet_l.columns: worksheet_l.column_dimensions[col[0].column_letter].width = min(max((len(str(cell.value)) for cell in col), default=0) + 2, 50) 
            
            # Add Not Attending Tab
            if not_attending_list:
                na_data = df_merged_full[df_merged_full['Official Name'].isin(not_attending_list)]
                na_export_df = na_data[['Student ID', 'Email', 'Official Name', 'Which Connections class are you in? ']].copy()
                na_export_df.rename(columns={
                    'Official Name': 'Student', 
                    'Which Connections class are you in? ': 'Class Group'
                }, inplace=True)
                
                na_export_df.to_excel(writer, sheet_name="Not Attending", index=False)
                worksheet_na = writer.sheets["Not Attending"]
                for col in worksheet_na.columns:
                    worksheet_na.column_dimensions[col[0].column_letter].width = max((len(str(cell.value)) for cell in col), default=0) + 2

        output.seek(0)
        st.markdown("---")
        st.download_button(label="📥 Download Final Excel File", data=output, file_name="Camp_Allocations_Final.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    else:
        st.info("👈 Please upload both CSV files in the sidebar to begin.")


# =========================================================================================
# ========================= PAGE 2: FINAL PUBLIC ROSTER BUILDER ===========================
# =========================================================================================
elif page == "📋 Final Roster & Leader Builder":
    st.title("📋 Final Roster & Leader Builder")

    roster_tab_y8, roster_tab_y9 = st.tabs(["🏕️ Year 8 Camps", "🏔️ Year 9 Camps"])

    # ── Year 8 Roster ─────────────────────────────────────────────────────────────────────
    with roster_tab_y8:
        st.subheader("Year 8 — Final Public Roster & Leader Builder")
        st.info("Upload your manually tweaked `Camp_Allocations_Final.xlsx` file here. The app will visually scan where you dragged the students in Excel, build a privacy-safe public roster, and completely regenerate an up-to-date Leader Overview for your staff.")

        if df_merged_full is None:
            st.warning("⚠️ Please upload the Y8 Form Responses and Y8 Student List CSVs in the sidebar (Step 1) to use this tool.")
        else:
            uploaded_xlsx = st.file_uploader("Upload Modified Excel File (.xlsx)", type=["xlsx"])
        
            if uploaded_xlsx:
                xls = pd.ExcelFile(uploaded_xlsx)
                sheet_names = xls.sheet_names
            
                master_assignments = []
                not_attending_ids = []
            
                # Create a set of all genuinely known Student IDs from the master list (to filter out random teacher notes!)
                valid_student_ids = set(df_merged_full['Student ID'].dropna().astype(str).str.replace(r'\.0$', '', regex=True).str.strip().tolist())
                valid_student_ids.discard('nan')
                valid_student_ids.discard('N/A')
                valid_student_ids.discard('')
            
                # 1. Capture the Not Attending list directly from the Excel file
                if "Not Attending" in sheet_names:
                    df_na = pd.read_excel(xls, sheet_name="Not Attending")
                    if 'Student ID' in df_na.columns:
                        not_attending_ids = df_na['Student ID'].dropna().astype(str).str.replace(r'\.0$', '', regex=True).tolist()
            
                # 2. Visually scan the Class tabs to see where the user placed the students
                for sheet in sheet_names:
                    if sheet in ["Leader Overview", "Not Attending"]:
                        continue # Skip non-class sheets
                
                    df_class = pd.read_excel(xls, sheet_name=sheet)
                    if 'Student ID' in df_class.columns:
                        current_camp = "Freycinet" # Default top section
                    
                        for _, row in df_class.iterrows():
                            stud_val = str(row.get('Student', '')).strip()
                            st_id = str(row.get('Student ID', '')).replace('.0', '').strip()
                        
                            # When the scanner hits the visual separator, flip the active camp
                            if 'BAY OF FIRES' in stud_val.upper() or 'BAY OF FIRES' in st_id.upper():
                                current_camp = "Bay of Fires"
                                continue
                            
                            # Save the assignment - ONLY if it's a verified valid Student ID (ignores ALL teacher notes/blanks perfectly)
                            if st_id in valid_student_ids and st_id not in not_attending_ids:
                                master_assignments.append({
                                    'Class': str(sheet).strip(),
                                    'Student ID': st_id,
                                    'Assigned Camp': current_camp
                                })
            
                if not master_assignments:
                    st.error("Could not find any valid student assignments in the uploaded Excel file.")
                else:
                    df_assignments = pd.DataFrame(master_assignments)
                    st.success(f"✅ Successfully read data for {len(df_assignments)} students across {df_assignments['Class'].nunique()} classes!")
                
                    # --- A. BUILD PUBLIC ROSTER (DATAFRAME FOR PREVIEW) ---
                    df_pub = pd.merge(df_assignments, df_stud_pub[['Student ID', 'Preferred name', 'Surname']], on='Student ID', how='left')
                    df_pub = df_pub[['Class', 'Preferred name', 'Surname', 'Assigned Camp']]
                    df_pub['Full Name'] = df_pub['Preferred name'].astype(str) + " " + df_pub['Surname'].astype(str)
                
                    # --- NEW FEATURE: GENERATE VISUALLY APPEALING PUBLIC EXCEL ROSTER ---
                    def generate_public_excel(df_roster):
                        wb = Workbook()
                        ws = wb.active
                        ws.title = "Camp Roster"

                        # Styles
                        title_font = Font(size=16, bold=True, color="FFFFFF")
                        f_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid") # Blue
                        b_fill = PatternFill(start_color="C0504D", end_color="C0504D", fill_type="solid") # Red
                        date_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid") # Dark Grey
                        class_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid") # Light Grey
                    
                        center_align = Alignment(horizontal="center", vertical="center")
                        left_align = Alignment(horizontal="left", vertical="center")
                        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                        unique_classes = sorted(df_roster['Class'].unique().tolist())
                    
                        # Dynamically match classes based on user's exact keywords
                        g1_names = ['rustin', 'stowe', 'preston']
                        g2_names = ['bracey', 'knight', 'tuke']
                    
                        g1_actual = [c for c in unique_classes if any(k in c.lower() for k in g1_names)]
                        g2_actual = [c for c in unique_classes if any(k in c.lower() for k in g2_names)]
                    
                        # Fallback if class names don't match those keywords at all
                        if not g1_actual and not g2_actual:
                            mid = len(unique_classes) // 2
                            g1_actual = unique_classes[:mid]
                            g2_actual = unique_classes[mid:]

                        def write_block(ws, start_row, date_text, classes):
                            if not classes: return start_row
                        
                            f_end_col = len(classes)
                            total_cols = f_end_col * 2
                        
                            # Row 1: Date Header
                            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=total_cols)
                            cell = ws.cell(row=start_row, column=1, value=date_text)
                            cell.font = title_font
                            cell.fill = date_fill
                            cell.alignment = center_align
                        
                            # Row 2: Camp Locations
                            ws.merge_cells(start_row=start_row+1, start_column=1, end_row=start_row+1, end_column=f_end_col)
                            cell_f = ws.cell(row=start_row+1, column=1, value="FREYCINET")
                            cell_f.font = Font(size=14, bold=True, color="FFFFFF")
                            cell_f.fill = f_fill
                            cell_f.alignment = center_align
                        
                            ws.merge_cells(start_row=start_row+1, start_column=f_end_col+1, end_row=start_row+1, end_column=total_cols)
                            cell_b = ws.cell(row=start_row+1, column=f_end_col+1, value="BAY OF FIRES")
                            cell_b.font = Font(size=14, bold=True, color="FFFFFF")
                            cell_b.fill = b_fill
                            cell_b.alignment = center_align
                        
                            # Row 3: Class Headers
                            for i, cls_name in enumerate(classes):
                                # Freycinet Side
                                c1 = ws.cell(row=start_row+2, column=i+1, value=cls_name)
                                c1.font = Font(bold=True)
                                c1.fill = class_fill
                                c1.alignment = center_align
                                c1.border = thin_border
                            
                                # Bay of Fires Side
                                c2 = ws.cell(row=start_row+2, column=f_end_col+i+1, value=cls_name)
                                c2.font = Font(bold=True)
                                c2.fill = class_fill
                                c2.alignment = center_align
                                c2.border = thin_border
                        
                            # Row 4+: Student Names
                            f_lists = []
                            b_lists = []
                            for cls_name in classes:
                                f_students = df_roster[(df_roster['Class'] == cls_name) & (df_roster['Assigned Camp'] == 'Freycinet')]['Full Name'].tolist()
                                b_students = df_roster[(df_roster['Class'] == cls_name) & (df_roster['Assigned Camp'] == 'Bay of Fires')]['Full Name'].tolist()
                                f_lists.append(f_students)
                                b_lists.append(b_students)
                        
                            max_len = max([len(l) for l in f_lists + b_lists] + [0])
                        
                            for row_offset in range(max_len):
                                for col_idx, f_list in enumerate(f_lists):
                                    val = f_list[row_offset] if row_offset < len(f_list) else ""
                                    cell = ws.cell(row=start_row+3+row_offset, column=col_idx+1, value=val)
                                    cell.alignment = left_align
                                    cell.border = thin_border
                                
                                for col_idx, b_list in enumerate(b_lists):
                                    val = b_list[row_offset] if row_offset < len(b_list) else ""
                                    cell = ws.cell(row=start_row+3+row_offset, column=f_end_col+col_idx+1, value=val)
                                    cell.alignment = left_align
                                    cell.border = thin_border
                                
                            # Return the next available row (with a 2-row buffer space)
                            return start_row + 3 + max_len + 2
                        
                        current_row = 2
                        if g1_actual:
                            current_row = write_block(ws, current_row, "Camp Dates: 11 - 15 May", g1_actual)
                        if g2_actual:
                            current_row = write_block(ws, current_row, "Camp Dates: 18 - 22 May", g2_actual)
                        
                        # Handle any extra stray classes just in case
                        other_actual = [c for c in unique_classes if c not in g1_actual and c not in g2_actual]
                        if other_actual:
                            for i in range(0, len(other_actual), 3):
                                chunk = other_actual[i:i+3]
                                current_row = write_block(ws, current_row, "Other Classes", chunk)
                            
                        # Set standard column widths
                        max_cols = (max(len(g1_actual), len(g2_actual), 1) * 2)
                        for col in range(1, max_cols + 1):
                            ws.column_dimensions[get_column_letter(col)].width = 25
                        
                        return wb

                    wb_public = generate_public_excel(df_pub)
                    output_pub = io.BytesIO()
                    wb_public.save(output_pub)
                    output_pub.seek(0)
                
                    # --- B. BUILD UPDATED LEADER OVERVIEW ---
                    leader_overview_data = []
                
                    for cls, class_group in df_assignments.groupby('Class'):
                        for camp_name in ['Freycinet', 'Bay of Fires']:
                            camp_students = class_group[class_group['Assigned Camp'] == camp_name]['Student ID'].tolist()
                            if not camp_students: continue
                        
                            # Add separator
                            leader_overview_data.append({
                                'Class': f"--- {cls} ---", 'Assigned Camp': f'--- {camp_name.upper()} ---', 'Student ID': '---', 'Email': '---', 'Student': '---------------------------',
                                'Responded': '---', 'Sea Kayak': '---', 'Coasteer': '---', 'MTB Interest': '---', 'Snorkel': '---',
                                'General Camping Skill': '---', 'Sleeping Outdoors/Bugs': '---', 'Swimming Confidence': '---', 'Bike Comfort': '---', 'Overnight Hike': '---',
                                'Reaction to Hardship': '---', 'Gear Independence': '---', 'Group Teamwork': '---'
                            })
                        
                            for st_id in camp_students:
                                stud_rows = df_merged_full[df_merged_full['Student ID'] == st_id]
                                if stud_rows.empty: continue
                                student_row = stud_rows.iloc[0]
                            
                                sk = student_row.get('How excited would you be to go sea kayaking?', 'N/A') if student_row['Responded'] else 'N/A'
                                coast = student_row.get('How excited would you be to go coasteering?', 'N/A') if student_row['Responded'] else 'N/A'
                                mtb_int = student_row.get('How excited would you be to go mountain bike riding on Blue (intermediate) flowy trails?', 'N/A') if student_row['Responded'] else 'N/A'
                                snork = student_row.get('How excited would you be to go snorkelling?', 'N/A') if student_row['Responded'] else 'N/A'
                                mtb_ab = student_row.get(mtb_ability_col, 'N/A') if mtb_ability_col and student_row['Responded'] else 'N/A'
                            
                                leader_overview_data.append({
                                    'Class': cls, 'Assigned Camp': camp_name, 
                                    'Student ID': st_id,
                                    'Email': student_row.get('Email', 'N/A'),
                                    'Student': student_row['Official Name'],
                                    'Responded': 'Yes' if student_row['Responded'] else 'No',
                                
                                    'Sea Kayak': sk if camp_name == 'Freycinet' else '---',
                                    'Coasteer': coast if camp_name == 'Freycinet' else '---',
                                    'MTB Interest': mtb_int if camp_name == 'Bay of Fires' else '---',
                                    'Snorkel': snork if camp_name == 'Bay of Fires' else '---',
                                
                                    'General Camping Skill': student_row.get('How confident are you in your general camping skills (like setting up a tent or cooking on a camp stove)?', 'N/A') if student_row['Responded'] else 'N/A',
                                    'Sleeping Outdoors/Bugs': student_row.get('How do you feel about sleeping outdoors, using long-drop toilets, and being around dirt and bugs for a few days?', 'N/A') if student_row['Responded'] else 'N/A',
                                    'Swimming Confidence': student_row.get('How confident are you swimming in deep water where you cannot touch the bottom?', 'N/A') if student_row['Responded'] else 'N/A',
                                    'Bike Comfort': mtb_ab,
                                    'Overnight Hike': student_row.get('How excited are you to go on an overnight hike carrying your gear?', 'N/A') if student_row['Responded'] else 'N/A',
                                    'Reaction to Hardship': student_row.get('When a physical activity gets tiring, difficult, or the weather turns bad, how do you usually react?', 'N/A') if student_row['Responded'] else 'N/A',
                                    'Gear Independence': student_row.get('How good are you at organising your own gear, packing your bag, and keeping yourself comfortable and safe without teachers or parents reminding you?', 'N/A') if student_row['Responded'] else 'N/A',
                                    'Group Teamwork': student_row.get('How confident are you in helping your group work together, joining in discussions, and making classmates feel included?', 'N/A') if student_row['Responded'] else 'N/A'
                                })
                            
                        leader_overview_data.append({
                            'Class': "", 'Assigned Camp': "", 'Student ID': "", 'Email': "", 'Student': "", 'Responded': "", 
                            'Sea Kayak': "", 'Coasteer': "", 'MTB Interest': "", 'Snorkel': "",
                            'General Camping Skill': "", 'Sleeping Outdoors/Bugs': "", 'Swimming Confidence': "", 'Bike Comfort': "", 'Overnight Hike': "",
                            'Reaction to Hardship': "", 'Gear Independence': "", 'Group Teamwork': ""
                        })
                    
                    df_leader_updated = pd.DataFrame(leader_overview_data)
                
                    # --- C. DISPLAY AND EXPORT ---
                    col1, col2 = st.columns(2)
                
                    with col1:
                        st.subheader("📋 Public Roster Processing Complete")
                        st.write("The simple flat list is shown below, but the exported Excel file is styled horizontally into groups (May 11-15 and May 18-22) for easy reading!")
                        st.dataframe(df_pub[['Class', 'Full Name', 'Assigned Camp']], hide_index=True, use_container_width=True)
                    
                        st.download_button(
                            label="📥 Download Stylized Public Roster (Excel)", 
                            data=output_pub, 
                            file_name="Final_Public_Camp_Roster.xlsx", 
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                            use_container_width=True
                        )
                    
                    with col2:
                        st.subheader("🛡️ Updated Leader Overview Preview")
                    
                        def highlight_leader_overview(row, df_ref):
                            colors = [''] * len(row)
                            if '---' in str(row.get('Student', '')):
                                return ['background-color: #d9d9d9; font-weight: bold; text-align: center'] * len(row)
                        
                            cols_to_check = ['General Camping Skill', 'Sleeping Outdoors/Bugs', 'Swimming Confidence', 'Bike Comfort', 'Overnight Hike', 'Reaction to Hardship', 'Gear Independence', 'Group Teamwork', 'Sea Kayak', 'Coasteer', 'MTB Interest', 'Snorkel']
                            for col_name in cols_to_check:
                                if col_name in df_ref.columns:
                                    col_idx = df_ref.columns.get_loc(col_name)
                                    val = str(row[col_name]).strip()
                                    match = re.search(r'^\s*([1-5])(?:\D|$)', val)
                                    if match:
                                        num = int(match.group(1))
                                        if num in [1, 2]: colors[col_idx] = 'background-color: #a83232; color: white; font-weight: bold'
                                        elif num in [3, 4]: colors[col_idx] = 'background-color: #ff9933; color: black; font-weight: bold'
                                        elif num == 5: colors[col_idx] = 'background-color: #ffcc00; color: black; font-weight: bold'
                            return colors

                        styled_leaders = df_leader_updated.style.apply(highlight_leader_overview, df_ref=df_leader_updated, axis=1)
                        st.dataframe(styled_leaders, use_container_width=True)
                    
                        output_ldr = io.BytesIO()
                        with pd.ExcelWriter(output_ldr, engine='openpyxl') as writer:
                            styled_leaders.to_excel(writer, sheet_name="Updated Leader Overview", index=False)
                            worksheet_l = writer.sheets["Updated Leader Overview"]
                            for col in worksheet_l.columns: worksheet_l.column_dimensions[col[0].column_letter].width = min(max((len(str(cell.value)) for cell in col), default=0) + 2, 50)
                        
                            if not_attending_ids:
                                na_data = df_merged_full[df_merged_full['Student ID'].isin(not_attending_ids)]
                                na_export_df = na_data[['Student ID', 'Email', 'Official Name', 'Which Connections class are you in? ']].copy()
                                na_export_df.rename(columns={'Official Name': 'Student', 'Which Connections class are you in? ': 'Class Group'}, inplace=True)
                                na_export_df.to_excel(writer, sheet_name="Not Attending", index=False)
                                worksheet_na = writer.sheets["Not Attending"]
                                for col in worksheet_na.columns: worksheet_na.column_dimensions[col[0].column_letter].width = max((len(str(cell.value)) for cell in col), default=0) + 2
                            
                        output_ldr.seek(0)
                        st.download_button(label="📥 Download Updated Leader Overview (Excel)", data=output_ldr, file_name="Updated_Leader_Overview.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    # ── Year 9 Roster ─────────────────────────────────────────────────────────────────────
    with roster_tab_y9:
        st.subheader("Year 9 — Final Public Roster Builder")
        st.info("Upload the `Y9_Journey_Groups.xlsx` file exported from the **Y9 Journey Groups** tool. The app will read each camp group sheet and build a clean public-facing roster organised by week and camp.")

        if not (y9_stud_file and y9_pref_file):
            st.warning("⚠️ Please upload the Y9 Student List and Y9 Preference Survey CSVs in the sidebar (Step 1) to enable student name lookups.")

        y9_roster_xlsx = st.file_uploader("Upload Y9_Journey_Groups.xlsx", type=["xlsx"], key="y9_roster_upload")

        if y9_roster_xlsx:
            # ── Camp key lookup ────────────────────────────────────────────────────────────
            _Y9R_CAMP_LABELS = {
                'MTB':        'Mountain to Sea MTB',
                'CC':         'Colossal Cliffs',
                'MM':         'Mersey & Mountains',
                'Explorer':   'Cradle to Coast Explorer',
                'Challenger': 'Cradle to Coast Challenger',
            }
            _Y9R_COLORS = {
                'MTB':        'FF2E7D32',
                'CC':         'FF1565C0',
                'MM':         'FF6A1B9A',
                'Explorer':   'FFE65100',
                'Challenger': 'FFB71C1C',
            }
            _WEEK_LABELS = {
                '1': 'Week 1: Unwin & Hodgkin (12–18 November)',
                '2': 'Week 2: Mather & Ransome (23–29 November)',
            }

            _xls_y9r = pd.ExcelFile(y9_roster_xlsx)
            _sheets_y9r = [s for s in _xls_y9r.sheet_names
                           if s not in ('Isolation Report', 'Not Attending')]

            # Parse sheets: name format  CAMPKEY[_A|_B]_W1 or CAMPKEY_W1
            _all_assignments = []
            for _sheet in _sheets_y9r:
                _parts = _sheet.split('_')
                _camp_key = _parts[0] if _parts else ''
                _week_num = _parts[-1].replace('W', '') if _parts else ''
                _subgroup = _parts[1] if len(_parts) == 3 else ''

                _df_sh = pd.read_excel(_xls_y9r, sheet_name=_sheet)
                for _, _row in _df_sh.iterrows():
                    _name = str(_row.get('Student', '')).strip()
                    if _name and _name.lower() not in ('nan', ''):
                        _all_assignments.append({
                            'Student':   _name,
                            'Camp Key':  _camp_key,
                            'Camp':      _Y9R_CAMP_LABELS.get(_camp_key, _camp_key),
                            'Group':     _subgroup if _subgroup else 'A',
                            'Week':      _week_num,
                            'Week Label': _WEEK_LABELS.get(_week_num, f'Week {_week_num}'),
                        })

            if not _all_assignments:
                st.error("Could not read any student assignments from the uploaded file. Check that it was exported from the Y9 Journey Groups tool.")
            else:
                _df_asgn = pd.DataFrame(_all_assignments)
                st.success(f"✅ Read {len(_df_asgn)} student assignments across {_df_asgn['Camp'].nunique()} camps and {_df_asgn['Week'].nunique()} weeks.")

                # ── Preview table ──────────────────────────────────────────────────────────
                for _wk in sorted(_df_asgn['Week'].unique()):
                    _wk_label = _WEEK_LABELS.get(str(_wk), f'Week {_wk}')
                    st.markdown(f"### 📅 {_wk_label}")
                    _wk_df = _df_asgn[_df_asgn['Week'] == _wk]
                    for _ck in [k for k in _Y9R_CAMP_LABELS if k in _wk_df['Camp Key'].values]:
                        _camp_df = _wk_df[_wk_df['Camp Key'] == _ck]
                        _has_b = 'B' in _camp_df['Group'].values
                        _camp_label = _Y9R_CAMP_LABELS.get(_ck, _ck)
                        if _has_b:
                            _ca = _camp_df[_camp_df['Group'] == 'A']['Student'].tolist()
                            _cb = _camp_df[_camp_df['Group'] == 'B']['Student'].tolist()
                            _col_a, _col_b = st.columns(2)
                            with _col_a:
                                st.markdown(f"**{_camp_label} — Group A** ({len(_ca)} students)")
                                st.dataframe(pd.DataFrame({'Student': _ca}), hide_index=True, use_container_width=True)
                            with _col_b:
                                st.markdown(f"**{_camp_label} — Group B** ({len(_cb)} students)")
                                st.dataframe(pd.DataFrame({'Student': _cb}), hide_index=True, use_container_width=True)
                        else:
                            _sa = _camp_df['Student'].tolist()
                            st.markdown(f"**{_camp_label}** ({len(_sa)} students)")
                            st.dataframe(pd.DataFrame({'Student': _sa}), hide_index=True, use_container_width=True)

                # ── Excel public roster export ─────────────────────────────────────────────
                _y9r_out = io.BytesIO()
                _wb_y9r = Workbook()

                _header_font  = Font(size=13, bold=True, color="FFFFFF")
                _camp_font    = Font(size=11, bold=True, color="FFFFFF")
                _week_fill    = PatternFill(start_color="FF333333", end_color="FF333333", fill_type="solid")
                _hdr_fill     = PatternFill(start_color="FF444444", end_color="FF444444", fill_type="solid")
                _name_font    = Font(size=10)
                _center       = Alignment(horizontal="center", vertical="center")
                _left         = Alignment(horizontal="left",   vertical="center")
                _thin         = Border(left=Side(style='thin'), right=Side(style='thin'),
                                       top=Side(style='thin'), bottom=Side(style='thin'))

                for _wk in sorted(_df_asgn['Week'].unique()):
                    _ws_name = f"Week {_wk}"
                    _ws = _wb_y9r.create_sheet(title=_ws_name)
                    _wk_label = _WEEK_LABELS.get(str(_wk), f'Week {_wk}')

                    _wk_df = _df_asgn[_df_asgn['Week'] == _wk]
                    _active_camps = [k for k in _Y9R_CAMP_LABELS if k in _wk_df['Camp Key'].values]

                    # Build column layout: each camp (or camp A+B pair) gets columns
                    _col_headers = []  # list of (camp_key, group, col_index_1based)
                    _col_idx = 1
                    for _ck in _active_camps:
                        _cdf = _wk_df[_wk_df['Camp Key'] == _ck]
                        _has_b = 'B' in _cdf['Group'].values
                        if _has_b:
                            _col_headers.append((_ck, 'A', _col_idx)); _col_idx += 1
                            _col_headers.append((_ck, 'B', _col_idx)); _col_idx += 1
                        else:
                            _col_headers.append((_ck, 'A', _col_idx)); _col_idx += 1
                    _total_cols = _col_idx - 1

                    # Row 1: Week header spanning all columns
                    _ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=_total_cols)
                    _wh = _ws.cell(row=1, column=1, value=_wk_label)
                    _wh.font = _header_font; _wh.fill = _week_fill; _wh.alignment = _center

                    # Row 2: Camp headers
                    _prev_ck = None
                    _group_start = None
                    for _ck, _sg, _ci in _col_headers:
                        _c = _ws.cell(row=2, column=_ci)
                        _label = _Y9R_CAMP_LABELS.get(_ck, _ck)
                        _sg_suffix = f" — Group {_sg}" if any(h[0] == _ck and h[1] == 'B' for h in _col_headers) else ""
                        _c.value = _label + _sg_suffix
                        _c.font = Font(bold=True, color="FFFFFF", size=10)
                        _fill_color = _Y9R_COLORS.get(_ck, 'FF555555')
                        _c.fill = PatternFill(start_color=_fill_color, end_color=_fill_color, fill_type="solid")
                        _c.alignment = _center; _c.border = _thin
                        _ws.column_dimensions[get_column_letter(_ci)].width = 26

                    # Rows 3+: Student names
                    _col_lists = {}
                    for _ck, _sg, _ci in _col_headers:
                        _cdf = _wk_df[(_wk_df['Camp Key'] == _ck) & (_wk_df['Group'] == _sg)]
                        _col_lists[_ci] = _cdf['Student'].tolist()

                    _max_rows = max((len(v) for v in _col_lists.values()), default=0)
                    for _r in range(_max_rows):
                        for _ck, _sg, _ci in _col_headers:
                            _lst = _col_lists[_ci]
                            _val = _lst[_r] if _r < len(_lst) else ""
                            _cell = _ws.cell(row=3 + _r, column=_ci, value=_val)
                            _cell.font = _name_font; _cell.alignment = _left; _cell.border = _thin

                # Remove default empty sheet
                if 'Sheet' in _wb_y9r.sheetnames:
                    del _wb_y9r['Sheet']

                _wb_y9r.save(_y9r_out)
                _y9r_out.seek(0)

                st.markdown("---")
                st.download_button(
                    label="📥 Download Y9 Public Roster (Excel)",
                    data=_y9r_out,
                    file_name="Y9_Public_Roster.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

# =========================================================================================
# ========================= PAGE 3: Y9 JOURNEY GROUPS ====================================
# =========================================================================================
elif page == "🏔️ Y9 Journey Groups":
    st.title("🏔️ Year 9 Journey Group Sorter")

    # ── Camp Definitions ──────────────────────────────────────────────────────────────────
    Y9_CAMP_DEFS = {
        'MTB':        {'label': 'Mountain to Sea MTB',                                   'max_per': 13},
        'CC':         {'label': 'Colossal Cliffs',                                       'max_per': 12},
        'MM':         {'label': 'Mersey & Mountains',                                    'max_per': 12},
        'Explorer':   {'label': 'Cradle to Coast Explorer (more flexibility)',            'max_per': 16},
        'Challenger': {'label': 'Cradle to Coast Challenger (more physical challenge)',   'max_per': 10},
    }

    Y9_NAME_MAP = {
        'mountain to sea mtb':                           'MTB',
        'colossal cliffs':                               'CC',
        'mersey and mountains':                          'MM',
        'mersey & mountains':                            'MM',
        'cradle to coast explorer (more flexibility)':   'Explorer',
        'cradle to coast explorer':                      'Explorer',
        'cradle to coast challenger (more physical challenge)': 'Challenger',
        'cradle to coast challenger':                    'Challenger',
    }

    WEEK1_HOUSES = {'unwin', 'hodgkin'}
    WEEK2_HOUSES = {'mather', 'ransome'}
    HOUSE_WEEK = {**{h: 1 for h in WEEK1_HOUSES}, **{h: 2 for h in WEEK2_HOUSES}}

    # ── Default AI weights ────────────────────────────────────────────────────────────────
    Y9_DEFAULT_WEIGHTS = {
        'w_pref':     100,      # Preference rank score multiplier
        'w_friend':   6000,     # Friend pair in same camp reward
        'w_cap':      600000,   # Over-capacity penalty (per student over limit)
        'w_force':    1000000,  # Forced camp violated penalty
        'w_must':     1000000,  # Must-go-with split penalty
        'w_sep':      500000,   # Can't-go-with together penalty
        'sg_balance': 15,       # Sub-group size imbalance multiplier
        'sg_cap':     600000,   # Sub-group over-capacity penalty
        'sg_friend':  400,      # Sub-group friend pair reward
        # Gender balance weights (applied per camp/subgroup)
        'w_gender_sole':    150000,  # Exactly 1 of a gender in group (worst)
        'w_gender_duo':      30000,  # Exactly 2 of a gender in group (medium)
        'w_gender_allsame':   8000,  # All same gender (better than sole, worse than mixed)
    }

    # ── Session State ─────────────────────────────────────────────────────────────────────
    for _k, _v in [('y9_sep', []), ('y9_must', []), ('y9_force', {}),
                   ('y9_force_week', {}),
                   ('y9_na', []), ('y9_seed', 0), ('y9_results', {}), ('y9_states', {}),
                   ('y9_include_drafts', False)]:
        if _k not in st.session_state:
            st.session_state[_k] = _v
    if 'y9_weights' not in st.session_state:
        st.session_state.y9_weights = Y9_DEFAULT_WEIGHTS.copy()

    if not (responses_file and students_file):
        st.info("👈 Upload a **Preference Survey CSV** and **Student List CSV** in the sidebar (Step 1) to begin.")
    else:
        # ── Load & Clean ──────────────────────────────────────────────────────────────────
        responses_file.seek(0); students_file.seek(0)
        df_s = pd.read_csv(students_file)
        df_p = pd.read_csv(responses_file)

        # Student list
        df_s['Email'] = df_s['Email'].astype(str).str.strip().str.lower()
        df_s['Official Name'] = (df_s['Preferred name'].astype(str).str.strip()
                                 + " " + df_s['Surname'].astype(str).str.strip())
        df_s['Gender'] = df_s['Gender'].astype(str).str.strip().str.lower() if 'Gender' in df_s.columns else 'o'
        if 'Code' in df_s.columns:
            df_s['Student ID'] = df_s['Code'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        else:
            df_s['Student ID'] = "N/A"

        # House from student list (fallback for non-responders)
        _house_col_stud = next((c for c in df_s.columns if c.strip().lower() == 'house'), None)
        _rg_col = next((c for c in df_s.columns if 'rollgroup' in c.lower()), None)
        _letter_map = {'u': 'unwin', 'h': 'hodgkin', 'm': 'mather', 'r': 'ransome'}
        if _house_col_stud:
            df_s['House_stud'] = df_s[_house_col_stud].astype(str).str.strip().str.lower()
        elif _rg_col:
            df_s['House_stud'] = (df_s[_rg_col].astype(str).str.strip().str[-1]
                                  .str.lower().map(_letter_map).fillna(''))
        else:
            df_s['House_stud'] = ''

        # Prefs CSV
        df_p['Email address'] = df_p['Email address'].astype(str).str.strip().str.lower()

        # Rank columns: look for "[1]"…"[5]" in header
        _rank_cols = {}
        for _col in df_p.columns:
            for _i in range(1, 6):
                if f'[{_i}]' in _col and ('prefer' in _col.lower() or 'journey' in _col.lower()):
                    _rank_cols[_i] = _col; break

        # Friend columns
        _friend_cols = [c for c in df_p.columns if 'suggest one person' in c.lower()]

        # House column in prefs
        _house_col_pref = next(
            (c for c in df_p.columns if c.strip().lower() == 'which house are you in?'), None)
        if not _house_col_pref:
            _house_col_pref = next((c for c in df_p.columns if 'house' in c.lower()), None)

        # Skill columns
        _camping_col   = next((c for c in df_p.columns if 'general camping skills' in c.lower()), None)
        _sleeping_col  = next((c for c in df_p.columns if 'sleeping outdoors' in c.lower()), None)
        _overnight_col = next((c for c in df_p.columns if 'overnight hike' in c.lower()), None)
        _hardship_col  = next((c for c in df_p.columns if 'tiring' in c.lower() and 'weather' in c.lower()), None)
        _swim_col      = next((c for c in df_p.columns if 'swimming' in c.lower() and 'deep water' in c.lower()), None)
        _challenge_col = next((c for c in df_p.columns if 'physical challenge' in c.lower()), None)
        _ww_col        = next((c for c in df_p.columns if 'white water' in c.lower()), None)
        _teamwork_col  = next((c for c in df_p.columns if 'helping your group' in c.lower() or 'group teamwork' in c.lower()), None)

        # Merge student list with prefs
        df_y9 = pd.merge(
            df_s[['Email', 'Official Name', 'Gender', 'Student ID', 'House_stud']],
            df_p, left_on='Email', right_on='Email address', how='left'
        )
        df_y9['Responded'] = df_y9['Email address'].notna()

        # Determine house and week
        def _get_house(row):
            if row['Responded'] and _house_col_pref:
                h = str(row.get(_house_col_pref, '')).strip().lower()
                if h in HOUSE_WEEK: return h
            h = str(row.get('House_stud', '')).strip().lower()
            return h if h in HOUSE_WEEK else None

        df_y9['House'] = df_y9.apply(_get_house, axis=1)
        df_y9['Week'] = df_y9['House'].map(HOUSE_WEEK)

        # Parse camp preference rankings
        official_names_y9 = df_s['Official Name'].tolist()

        def _parse_prefs(row):
            prefs = {}
            if not row['Responded']: return prefs
            for rank_num, col in _rank_cols.items():
                val = str(row.get(col, '')).strip().lower()
                for frag, key in Y9_NAME_MAP.items():
                    if frag in val:
                        if key not in prefs: prefs[key] = rank_num
                        break
            return prefs

        df_y9['Camp Prefs'] = df_y9.apply(_parse_prefs, axis=1)

        # Friend name matching
        def _match_friend(raw):
            if not raw or not isinstance(raw, str): return None
            raw = raw.strip()
            if raw.lower() in ('', 'nan', 'none'): return None
            raw_l = raw.lower()
            for name in official_names_y9:
                if name.lower() == raw_l: return name
            raw_words = set(raw_l.split())
            for name in official_names_y9:
                if raw_words and raw_words.issubset(set(name.lower().split())): return name
            return None

        def _get_friends(row):
            friends = []
            if not row['Responded']: return friends
            for fc in _friend_cols:
                m = _match_friend(row.get(fc))
                if m and m != row['Official Name'] and m not in friends:
                    friends.append(m)
            return friends

        df_y9['Friends Requested'] = df_y9.apply(_get_friends, axis=1)

        # ── Summary metrics ───────────────────────────────────────────────────────────────
        all_y9 = sorted(df_y9['Official Name'].dropna().unique().tolist())

        # ── SIDEBAR (only visible on this page) ───────────────────────────────────────────
        st.sidebar.markdown("---")

        # ── Load Rules & Weights ──────────────────────────────────────────────────────────
        st.sidebar.header("📂 Load Rules & Weights")
        _y9_json_upload = st.sidebar.file_uploader("Load Saved Y9 JSON", type=["json"], key="y9_json_upload")
        if _y9_json_upload is not None:
            if st.session_state.get('y9_last_loaded') != _y9_json_upload.file_id:
                try:
                    _loaded = json.load(_y9_json_upload)
                    st.session_state.y9_sep        = [tuple(x) for x in _loaded.get('sep',   [])]
                    st.session_state.y9_must       = [tuple(x) for x in _loaded.get('must',  [])]
                    st.session_state.y9_force      = _loaded.get('force', {})
                    st.session_state.y9_force_week = _loaded.get('force_week', {})
                    _l_na = [s for s in _loaded.get('na', []) if s in all_y9]
                    st.session_state.y9_na = _l_na
                    if 'weights' in _loaded:
                        st.session_state.y9_weights.update(_loaded['weights'])
                    st.session_state.y9_include_drafts = _loaded.get('include_drafts', False)
                    st.session_state.y9_last_loaded = _y9_json_upload.file_id
                    st.sidebar.success("✅ Y9 settings loaded!")
                    st.rerun()
                except Exception:
                    st.sidebar.error("❌ Could not parse the JSON file.")

        st.sidebar.markdown("---")

        # ── Not Attending ─────────────────────────────────────────────────────────────────
        st.sidebar.header("🏔️ Y9: Not Attending")
        _y9_na = st.sidebar.multiselect(
            "Students not attending:", options=all_y9,
            default=st.session_state.y9_na, key="y9_na_ms")
        st.session_state.y9_na = _y9_na

        y9_include_drafts = st.sidebar.checkbox(
            "✏️ Draft in students who haven't responded (shown in grey)",
            value=st.session_state.y9_include_drafts, key="y9_drafts_cb")
        st.session_state.y9_include_drafts = y9_include_drafts
        if y9_include_drafts:
            st.sidebar.caption("Non-responding students will be slotted into groups and shown with a grey background.")

        df_y9_act = df_y9[~df_y9['Official Name'].isin(_y9_na)].copy()
        attending_y9 = sorted(df_y9_act['Official Name'].dropna().unique().tolist())

        # Optimisation pool: responders only (unless draft mode is on)
        if y9_include_drafts:
            df_y9_opt = df_y9_act.copy()
        else:
            df_y9_opt = df_y9_act[df_y9_act['Responded']].copy()

        # Force camp
        st.sidebar.header("📍 Y9: Force Camp")
        _y9_fs = st.sidebar.selectbox("Student", [""] + attending_y9, key="y9_frc_s")
        _y9_fc = st.sidebar.selectbox(
            "Camp", [''] + list(Y9_CAMP_DEFS.keys()),
            format_func=lambda k: Y9_CAMP_DEFS[k]['label'] if k else "-- select --",
            key="y9_frc_c")
        if st.sidebar.button("Add Force Rule", key="y9_add_force"):
            if _y9_fs and _y9_fc:
                st.session_state.y9_force[_y9_fs] = _y9_fc; st.rerun()
        if st.session_state.y9_force:
            st.sidebar.write("**Active Force Rules:**")
            for _s, _c in list(st.session_state.y9_force.items()):
                _c1, _c2 = st.sidebar.columns([5, 1])
                _c1.write(f"- {_s} ➡️ {Y9_CAMP_DEFS.get(_c, {}).get('label', _c)[:22]}")
                if _c2.button("❌", key=f"y9_df_{_s}"):
                    del st.session_state.y9_force[_s]; st.rerun()
            if st.sidebar.button("Clear Force Rules", key="y9_clr_force"):
                st.session_state.y9_force = {}; st.rerun()

        # Force week (opposite-week override)
        st.sidebar.header("🔀 Y9: Force to Opposite Week")
        st.sidebar.caption("Use this when a student must go in the opposite week to their house. E.g. a Unwin/Hodgkin student going with Mather/Ransome.")
        _y9_fw_s = st.sidebar.selectbox("Student", [""] + attending_y9, key="y9_fwk_s")
        _y9_fw_w = st.sidebar.selectbox("Send to Week", ["", 1, 2],
            format_func=lambda w: "-- select --" if w == "" else (
                "Week 1 (Unwin & Hodgkin)" if w == 1 else "Week 2 (Mather & Ransome)"),
            key="y9_fwk_w")
        if st.sidebar.button("Add Force Week Rule", key="y9_add_fwk"):
            if _y9_fw_s and _y9_fw_w != "":
                st.session_state.y9_force_week[_y9_fw_s] = int(_y9_fw_w); st.rerun()
        if st.session_state.y9_force_week:
            st.sidebar.write("**Active Force Week Rules:**")
            for _s, _w in list(st.session_state.y9_force_week.items()):
                _c1, _c2 = st.sidebar.columns([5, 1])
                _wlbl = "Wk 1" if _w == 1 else "Wk 2"
                _c1.write(f"- {_s} ➡️ {_wlbl}")
                if _c2.button("❌", key=f"y9_dfw_{_s}"):
                    del st.session_state.y9_force_week[_s]; st.rerun()
            if st.sidebar.button("Clear Force Week Rules", key="y9_clr_fwk"):
                st.session_state.y9_force_week = {}; st.rerun()

        # Must go with
        st.sidebar.header("🤝 Y9: Must Go With")
        _cm1, _cm2 = st.sidebar.columns(2)
        _y9_m1 = _cm1.selectbox("A", [""] + attending_y9, key="y9_m1")
        _y9_m2 = _cm2.selectbox("B", [""] + attending_y9, key="y9_m2")
        if st.sidebar.button("Add Must-Go Rule", key="y9_add_must"):
            if _y9_m1 and _y9_m2 and _y9_m1 != _y9_m2:
                st.session_state.y9_must.append((_y9_m1, _y9_m2)); st.rerun()
        if st.session_state.y9_must:
            st.sidebar.write("**Active Must-Go Rules:**")
            for _i, _pair in enumerate(st.session_state.y9_must):
                _c1, _c2 = st.sidebar.columns([5, 1])
                _c1.write(f"- {_pair[0]} & {_pair[1]}")
                if _c2.button("❌", key=f"y9_dm_{_i}"):
                    st.session_state.y9_must.pop(_i); st.rerun()
            if st.sidebar.button("Clear Must-Go Rules", key="y9_clr_must"):
                st.session_state.y9_must = []; st.rerun()

        # Can't go with
        st.sidebar.header("⚠️ Y9: Can't Go With")
        _cs1, _cs2 = st.sidebar.columns(2)
        _y9_s1 = _cs1.selectbox("S1", [""] + attending_y9, key="y9_s1")
        _y9_s2 = _cs2.selectbox("S2", [""] + attending_y9, key="y9_s2")
        if st.sidebar.button("Add Can't-Go Rule", key="y9_add_sep"):
            if _y9_s1 and _y9_s2 and _y9_s1 != _y9_s2:
                st.session_state.y9_sep.append((_y9_s1, _y9_s2)); st.rerun()
        if st.session_state.y9_sep:
            st.sidebar.write("**Active Can't-Go Rules:**")
            for _i, _pair in enumerate(st.session_state.y9_sep):
                _c1, _c2 = st.sidebar.columns([5, 1])
                _c1.write(f"- {_pair[0]} ≠ {_pair[1]}")
                if _c2.button("❌", key=f"y9_ds_{_i}"):
                    st.session_state.y9_sep.pop(_i); st.rerun()
            if st.sidebar.button("Clear Can't-Go Rules", key="y9_clr_sep"):
                st.session_state.y9_sep = []; st.rerun()

        st.sidebar.markdown("---")

        # ── Advanced AI Weights ───────────────────────────────────────────────────────────
        with st.sidebar.expander("⚖️ Advanced AI Weights"):
            st.write("Tune how strongly the AI rewards or penalises each factor.")
            _yw = st.session_state.y9_weights
            _yw['w_pref']     = st.number_input("Preference Score Weight",       value=_yw['w_pref'],     step=10,    key="yw_pref")
            _yw['w_friend']   = st.number_input("Friend Pair Reward",            value=_yw['w_friend'],   step=500,   key="yw_friend")
            _yw['w_cap']      = st.number_input("Over-Capacity Penalty",         value=_yw['w_cap'],      step=10000, key="yw_cap")
            _yw['w_force']    = st.number_input("Forced Camp Penalty",           value=_yw['w_force'],    step=10000, key="yw_force")
            _yw['w_must']     = st.number_input("Must-Go-With Penalty",          value=_yw['w_must'],     step=10000, key="yw_must")
            _yw['w_sep']      = st.number_input("Can't-Go-With Penalty",         value=_yw['w_sep'],      step=10000, key="yw_sep")
            _yw['sg_friend']  = st.number_input("Sub-group Friend Reward",       value=_yw['sg_friend'],  step=50,    key="yw_sg_friend")
            st.markdown("**Gender Balance**")
            _yw['w_gender_sole']    = st.number_input("Sole gender in group (1 of a gender) — Penalty",    value=_yw.get('w_gender_sole',    150000), step=10000, key="yw_g_sole")
            _yw['w_gender_duo']     = st.number_input("Duo gender in group (2 of a gender) — Penalty",     value=_yw.get('w_gender_duo',      30000), step=5000,  key="yw_g_duo")
            _yw['w_gender_allsame'] = st.number_input("All same gender in group — Penalty",                value=_yw.get('w_gender_allsame',   8000), step=1000,  key="yw_g_allsame")
            if st.button("↩️ Reset to Defaults", key="y9_reset_wts"):
                st.session_state.y9_weights = Y9_DEFAULT_WEIGHTS.copy(); st.rerun()

        st.sidebar.markdown("---")

        # ── Save Your Setup ───────────────────────────────────────────────────────────────
        st.sidebar.header("💾 Save Your Setup")
        _na_details = []
        for _na_name in st.session_state.y9_na:
            _na_row = df_y9[df_y9['Official Name'] == _na_name]
            if not _na_row.empty:
                _nr = _na_row.iloc[0]
                _na_details.append({
                    'name': _na_name,
                    'student_id': str(_nr.get('Student ID', '')),
                    'house': str(_nr.get('House_stud', '')),
                    'email': str(_nr.get('Email', '')),
                })
        _y9_save = {
            'sep':        [list(p) for p in st.session_state.y9_sep],
            'must':       [list(p) for p in st.session_state.y9_must],
            'force':      st.session_state.y9_force,
            'force_week': st.session_state.y9_force_week,
            'na':         st.session_state.y9_na,
            'na_details': _na_details,
            'include_drafts': st.session_state.y9_include_drafts,
            'weights':    st.session_state.y9_weights,
        }
        st.sidebar.download_button(
            label="📥 Download Y9 Rules & Weights (JSON)",
            data=json.dumps(_y9_save, indent=2),
            file_name="y9_camp_rules.json",
            mime="application/json",
            use_container_width=True,
            key="y9_save_btn"
        )

        # ── MAIN CONTROLS ─────────────────────────────────────────────────────────────────
        st.markdown("---")
        _col_find, _col_ctrl = st.columns([2, 1])
        with _col_find:
            st.subheader("🔍 Find a Student")
            y9_find = st.selectbox("Search:", [""] + attending_y9, key="y9_find_s")
        with _col_ctrl:
            st.subheader("⚙️ Controls")
            y9_depth = st.select_slider(
                "Search Depth:", ["Fast (1-2s)", "Standard (5-10s)", "Deep Search (20-40s)"],
                value="Standard (5-10s)", key="y9_depth")
            if st.button("🔄 Regenerate Y9 Groups", use_container_width=True, key="y9_regen"):
                st.session_state.y9_seed += 1; st.rerun()

        y9_find_result = st.empty()
        st.markdown("---")

        # ── WEEK SUMMARY ──────────────────────────────────────────────────────────────────
        _force_week_rules = st.session_state.get('y9_force_week', {})
        # Compute effective week counts after force-week overrides
        def _effective_week(row):
            nm = row['Official Name']
            if nm in _force_week_rules:
                return _force_week_rules[nm]
            return row['Week']
        df_y9_act['_eff_week'] = df_y9_act.apply(_effective_week, axis=1)
        _w1_n = len(df_y9_act[df_y9_act['_eff_week'] == 1])
        _w2_n = len(df_y9_act[df_y9_act['_eff_week'] == 2])
        _wX_n = len(df_y9_act[df_y9_act['_eff_week'].isna()])
        _fw_n = len([s for s in _force_week_rules if s in attending_y9])
        _mc1, _mc2, _mc3, _mc4, _mc5 = st.columns(5)
        _mc1.metric("Total Y9", len(df_y9_act))
        _mc2.metric("Week 1 (Unwin + Hodgkin)", _w1_n)
        _mc3.metric("Week 2 (Mather + Ransome)", _w2_n)
        _mc4.metric("⚠️ No Week Assigned", _wX_n)
        _mc5.metric("🔀 Force-Week Overrides", _fw_n)
        if _wX_n:
            with st.expander("⚠️ Students with no house/week assigned — check their data"):
                _unassigned = df_y9_act[df_y9_act['Week'].isna()][['Student ID', 'Official Name', 'Responded']]
                st.dataframe(_unassigned.rename(columns={'Official Name': 'Student'}), hide_index=True)

        # ── ALGORITHM HELPERS ─────────────────────────────────────────────────────────────

        def _determine_config(n_students, pref_data_dict, week_num):
            """Return {camp_key: n_instances} obeying per-week camp count limits.

            Rules (same for both weeks unless noted):
              • MTB:  exactly 1
              • CC:   1 or 2
              • MM:   1 or 2
              • Cradle (Explorer + Challenger combined): exactly 2 instances total
                  — could be 2×Explorer, 2×Challenger, or 1 of each
              • Week 1 total: ≤ 7 camps   Week 2 total: ≤ 6 camps
            """
            MAX_CAMPS = {1: 7, 2: 6}.get(week_num, 7)

            # Count demand for each cradle type
            exp_d = sum(1 for p in pref_data_dict.values() if p.get('Explorer',   5) <= 2)
            cha_d = sum(1 for p in pref_data_dict.values() if p.get('Challenger', 5) <= 2)

            # Cradle split: always 2 instances total; split by demand
            if exp_d >= cha_d:
                if exp_d == 0 and cha_d == 0:
                    cradle_config = {'Explorer': 1, 'Challenger': 1}   # no preference data — split evenly
                elif cha_d == 0:
                    cradle_config = {'Explorer': 2}
                else:
                    cradle_config = {'Explorer': 1, 'Challenger': 1}
            else:
                if exp_d == 0:
                    cradle_config = {'Challenger': 2}
                else:
                    cradle_config = {'Explorer': 1, 'Challenger': 1}

            # Start with the fixed minimum: MTB×1, CC×1, MM×1, plus cradle
            config = {'MTB': 1, 'CC': 1, 'MM': 1, **cradle_config}
            total = sum(config.values())   # 4 or 5 depending on cradle split

            # Count demand for CC and MM to decide whether to run a second instance
            cc_d = sum(1 for p in pref_data_dict.values() if p.get('CC', 5) <= 2)
            mm_d = sum(1 for p in pref_data_dict.values() if p.get('MM', 5) <= 2)

            # Add a second MM first (if demand warrants and we have room)
            if total < MAX_CAMPS and mm_d >= cc_d:
                config['MM'] = 2; total += 1
            # Add a second CC (if room)
            if total < MAX_CAMPS:
                config['CC'] = 2; total += 1
            # If MM wasn't doubled yet, try again
            if total < MAX_CAMPS and config.get('MM', 1) == 1:
                config['MM'] = 2; total += 1

            return config

        def _get_caps(config):
            return {k: Y9_CAMP_DEFS[k]['max_per'] * v for k, v in config.items()}

        def _optimize_assignment(week_df, config, caps, friend_reqs, forced,
                                 must_pairs, sep_pairs, seed_str, depth, gender_lookup=None):
            """Phase 1: assign every student to a camp TYPE via hill-climbing."""
            active = list(config.keys())
            students = list(week_df['Official Name'])
            pref_data = dict(zip(week_df['Official Name'], week_df['Camp Prefs']))

            _yw = st.session_state.y9_weights
            W_PREF   = _yw.get('w_pref',   100)
            W_FRIEND = _yw.get('w_friend',  6000)
            W_CAP    = _yw.get('w_cap',     600000)
            W_FORCE  = _yw.get('w_force',   1000000)
            W_MUST   = _yw.get('w_must',    1000000)
            W_SEP    = _yw.get('w_sep',     500000)
            W_G_SOLE    = _yw.get('w_gender_sole',    150000)
            W_G_DUO     = _yw.get('w_gender_duo',      30000)
            W_G_ALLSAME = _yw.get('w_gender_allsame',   8000)

            def _score(asgn):
                s = 0
                counts = {c: 0 for c in active}
                for _c in asgn.values(): counts[_c] += 1
                for _c, _n in counts.items():
                    if _n > caps[_c]: s -= W_CAP * (_n - caps[_c])
                for _st, _camp in asgn.items():
                    prefs = pref_data.get(_st, {})
                    s += (6 - prefs.get(_camp, 3)) * W_PREF
                    for _f in friend_reqs.get(_st, []):
                        if asgn.get(_f) == _camp: s += W_FRIEND
                    if _st in forced and forced[_st] in active:
                        if forced[_st] != _camp: s -= W_FORCE
                for _a, _b in must_pairs:
                    if _a in asgn and _b in asgn and asgn[_a] != asgn[_b]: s -= W_MUST
                for _a, _b in sep_pairs:
                    if _a in asgn and _b in asgn and asgn[_a] == asgn[_b]: s -= W_SEP
                # Gender balance per camp
                if gender_lookup:
                    for _camp in active:
                        _members = [st for st, c in asgn.items() if c == _camp]
                        _m = sum(1 for st in _members if gender_lookup.get(st, 'o') == 'm')
                        _f = sum(1 for st in _members if gender_lookup.get(st, 'o') == 'f')
                        # All same gender (no mix): light penalty
                        if _m == 0 and _f > 0: s -= W_G_ALLSAME
                        elif _f == 0 and _m > 0: s -= W_G_ALLSAME
                        else:
                            # Mixed: penalise lone or duo of either gender
                            if _m == 1: s -= W_G_SOLE
                            elif _m == 2: s -= W_G_DUO
                            if _f == 1: s -= W_G_SOLE
                            elif _f == 2: s -= W_G_DUO
                return s

            if "Fast" in depth:   restarts, iters = 30,  2000
            elif "Standard" in depth: restarts, iters = 100, 5000
            else:                 restarts, iters = 300, 10000

            random.seed(seed_str)
            best_asgn, best_s = None, -float('inf')

            for _ in range(restarts):
                curr = {}
                load = {c: 0 for c in active}
                shuffled = students.copy(); random.shuffle(shuffled)
                for _st in shuffled:
                    if _st in forced and forced[_st] in active:
                        chosen = forced[_st]
                    else:
                        prefs = pref_data.get(_st, {})
                        sorted_camps = sorted(active, key=lambda c: prefs.get(c, 3))
                        chosen = sorted_camps[0]
                        for _c in sorted_camps:
                            if load[_c] < caps[_c]: chosen = _c; break
                    curr[_st] = chosen; load[chosen] += 1

                curr_s = _score(curr)
                for _ in range(iters):
                    new = dict(curr)
                    if random.random() < 0.7 and len(students) >= 2:
                        _s1, _s2 = random.sample(students, 2)
                        new[_s1], new[_s2] = curr[_s2], curr[_s1]
                    else:
                        _st = random.choice(students)
                        others = [c for c in active if c != curr[_st]]
                        if others: new[_st] = random.choice(others)
                    ns = _score(new)
                    if ns > curr_s: curr = new; curr_s = ns

                if curr_s > best_s: best_s = curr_s; best_asgn = dict(curr)

            return best_asgn

        def _split_subgroups(camp_students, max_per, friend_reqs, seed_str, gender_lookup=None):
            """Phase 2: split a camp's student list into sub-groups A and B."""
            students = list(camp_students)
            n = len(students)
            if n == 0: return [], []
            if n <= max_per: return students, []

            _yw = st.session_state.y9_weights
            SG_BAL    = _yw.get('sg_balance', 15)
            SG_CAP    = _yw.get('sg_cap',     600000)
            SG_FRIEND = _yw.get('sg_friend',  400)
            SG_G_SOLE    = _yw.get('w_gender_sole',    150000) // 2
            SG_G_DUO     = _yw.get('w_gender_duo',      30000) // 2
            SG_G_ALLSAME = _yw.get('w_gender_allsame',   8000) // 2

            def _sg_score(a, b):
                s = -(abs(len(a) - len(b)) ** 2) * SG_BAL
                if len(a) > max_per: s -= SG_CAP * (len(a) - max_per)
                if len(b) > max_per: s -= SG_CAP * (len(b) - max_per)
                for _st in a:
                    for _f in friend_reqs.get(_st, []):
                        if _f in a: s += SG_FRIEND
                for _st in b:
                    for _f in friend_reqs.get(_st, []):
                        if _f in b: s += SG_FRIEND
                # Gender balance per subgroup
                if gender_lookup:
                    for _grp in [a, b]:
                        if len(_grp) == 0: continue
                        _m = sum(1 for st in _grp if gender_lookup.get(st, 'o') == 'm')
                        _f = sum(1 for st in _grp if gender_lookup.get(st, 'o') == 'f')
                        if _m == 0 and _f > 0: s -= SG_G_ALLSAME
                        elif _f == 0 and _m > 0: s -= SG_G_ALLSAME
                        else:
                            if _m == 1: s -= SG_G_SOLE
                            elif _m == 2: s -= SG_G_DUO
                            if _f == 1: s -= SG_G_SOLE
                            elif _f == 2: s -= SG_G_DUO
                return s

            random.seed(seed_str)
            target = n // 2 + n % 2
            shuffled = students.copy(); random.shuffle(shuffled)
            best_a = set(shuffled[:target]); best_b = set(shuffled[target:])
            best_s = _sg_score(best_a, best_b)

            for _ in range(3000):
                na, nb = set(best_a), set(best_b)
                if na and nb and random.random() < 0.8:
                    _s1 = random.choice(sorted(na)); _s2 = random.choice(sorted(nb))
                    na.discard(_s1); na.add(_s2); nb.discard(_s2); nb.add(_s1)
                elif random.random() < 0.5 and na:
                    _st = random.choice(sorted(na)); na.discard(_st); nb.add(_st)
                elif nb:
                    _st = random.choice(sorted(nb)); nb.discard(_st); na.add(_st)
                ns = _sg_score(na, nb)
                if ns > best_s: best_s = ns; best_a, best_b = na, nb

            return sorted(list(best_a)), sorted(list(best_b))

        # ── PER-WEEK OPTIMIZATION ─────────────────────────────────────────────────────────
        all_week_results = {}  # {week_num: {camp_key: {'A': [...], 'B': [...]}}}
        all_week_configs  = {}
        all_friend_reqs   = {}

        for _wk, _wk_label in [(1, "Week 1 (Unwin & Hodgkin)"), (2, "Week 2 (Mather & Ransome)")]:
            _wdf = df_y9_opt[df_y9_opt['Week'] == _wk].copy()

            # Apply force-week overrides: students whose home week differs from _wk
            # but have been explicitly forced into this week are added here;
            # students forced out of this week are dropped.
            _force_week_rules = st.session_state.get('y9_force_week', {})
            _forced_into_wk   = [s for s, w in _force_week_rules.items() if w == _wk and s in attending_y9]
            _forced_out_of_wk = [s for s, w in _force_week_rules.items() if w != _wk and s in attending_y9]

            # Add forced-in students (may come from the other week's natural cohort)
            for _fi in _forced_into_wk:
                if _fi not in _wdf['Official Name'].values:
                    _fi_row = df_y9_act[df_y9_act['Official Name'] == _fi]
                    if not _fi_row.empty:
                        _wdf = pd.concat([_wdf, _fi_row], ignore_index=True)

            # Remove forced-out students
            _wdf = _wdf[~_wdf['Official Name'].isin(_forced_out_of_wk)]

            if _wdf.empty:
                all_week_results[_wk] = {}; all_week_configs[_wk] = {}; continue

            _week_students = set(_wdf['Official Name'])

            # Build friend requests (only within same week + must/sep rules applied)
            _wfr = {}
            for _, _row in _wdf.iterrows():
                _st = _row['Official Name']
                _fr = [f for f in _row['Friends Requested']
                       if f in _week_students and f not in _y9_na]
                for _a, _b in st.session_state.y9_must:
                    if _st == _a and _b in _week_students and _b not in _fr: _fr.append(_b)
                    if _st == _b and _a in _week_students and _a not in _fr: _fr.append(_a)
                for _a, _b in st.session_state.y9_sep:
                    if _st == _a and _b in _fr: _fr.remove(_b)
                    if _st == _b and _a in _fr: _fr.remove(_a)
                _wfr[_st] = _fr
            all_friend_reqs.update(_wfr)

            _pref_dict = dict(zip(_wdf['Official Name'], _wdf['Camp Prefs']))
            _config = _determine_config(len(_wdf), _pref_dict, _wk)
            _caps   = _get_caps(_config)
            all_week_configs[_wk] = _config

            _forced_wk = {k: v for k, v in st.session_state.y9_force.items() if k in _week_students}
            _sep_wk    = [(a, b) for a, b in st.session_state.y9_sep   if a in _week_students or b in _week_students]
            _must_wk   = [(a, b) for a, b in st.session_state.y9_must  if a in _week_students or b in _week_students]

            _state_key = str({
                'config': _config, 'na': sorted(_y9_na),
                'force': sorted(_forced_wk.items()),
                'force_week': sorted(_force_week_rules.items()),
                'sep':   sorted([tuple(sorted(p)) for p in _sep_wk]),
                'must':  sorted([tuple(sorted(p)) for p in _must_wk]),
                'seed': st.session_state.y9_seed, 'depth': y9_depth, 'n': len(_wdf),
                'drafts': y9_include_drafts,
                'g_sole': st.session_state.y9_weights.get('w_gender_sole', 150000),
                'g_duo':  st.session_state.y9_weights.get('w_gender_duo',   30000),
                'g_all':  st.session_state.y9_weights.get('w_gender_allsame', 8000),
            })

            if st.session_state.y9_states.get(_wk) != _state_key:
                st.session_state.y9_results.pop(_wk, None)
                st.session_state.y9_states[_wk] = _state_key

            if _wk not in st.session_state.y9_results:
                with st.spinner(f"⚙️ Optimising {_wk_label} — {len(_wdf)} students across {sum(_config.values())} camps…"):
                    _gender_lk = dict(zip(_wdf['Official Name'], _wdf['Gender']))
                    _asgn = _optimize_assignment(
                        _wdf, _config, _caps, _wfr,
                        _forced_wk, _must_wk, _sep_wk,
                        f"y9_w{_wk}_{st.session_state.y9_seed}", y9_depth,
                        gender_lookup=_gender_lk
                    )

                _week_groups = {}
                for _ck, _ni in _config.items():
                    _camp_list = sorted([s for s, c in _asgn.items() if c == _ck])
                    _mp = Y9_CAMP_DEFS[_ck]['max_per']
                    if _ni == 1:
                        _week_groups[_ck] = {'A': _camp_list, 'B': []}
                    else:
                        _ga, _gb = _split_subgroups(
                            _camp_list, _mp, _wfr,
                            f"split_{_ck}_w{_wk}_{st.session_state.y9_seed}",
                            gender_lookup=_gender_lk)
                        _week_groups[_ck] = {'A': _ga, 'B': _gb}

                st.session_state.y9_results[_wk] = _week_groups

            all_week_results[_wk] = st.session_state.y9_results[_wk]

        # ── BUILD FULL LOOKUP (student → week, camp, subgroup) ────────────────────────────
        full_lookup = {}
        for _wk, _wg in all_week_results.items():
            for _ck, _sgs in _wg.items():
                for _sg, _sgl in _sgs.items():
                    for _nm in _sgl:
                        full_lookup[_nm] = (_wk, _ck, _sg)

        # ── FIND STUDENT RESULT ───────────────────────────────────────────────────────────
        if y9_find and y9_find in full_lookup:
            _fw, _fc, _fsg = full_lookup[y9_find]
            _has_b = len(all_week_results.get(_fw, {}).get(_fc, {}).get('B', [])) > 0
            _sg_str = f" — Group {_fsg}" if _has_b else ""
            _wk_str = "Week 1: Unwin & Hodgkin" if _fw == 1 else "Week 2: Mather & Ransome"
            y9_find_result.success(
                f"🎯 **{y9_find}** → {_wk_str} | **{Y9_CAMP_DEFS[_fc]['label']}{_sg_str}**")
        elif y9_find:
            y9_find_result.warning(
                f"⚠️ {y9_find} was not placed — they may have no house/week assigned.")

        # ── COMPUTE ISOLATED STUDENTS ─────────────────────────────────────────────────────
        isolated_y9 = []
        _seen_iso = set()
        for _nm, (_fw, _fc, _fsg) in full_lookup.items():
            _friends = all_friend_reqs.get(_nm, [])
            if not _friends: continue
            for _f in _friends:
                if full_lookup.get(_f, (None, None, None))[:2] != (_fw, _fc):
                    if _nm not in _seen_iso:
                        _seen_iso.add(_nm)
                        _f_dest = full_lookup.get(_f)
                        isolated_y9.append({
                            'Week': _fw,
                            'Student': _nm,
                            'Assigned Camp': Y9_CAMP_DEFS[_fc]['label'],
                            'Friend Requested': _f,
                            'Friend\'s Camp': Y9_CAMP_DEFS[_f_dest[1]]['label'] if _f_dest else 'Not placed',
                        })
                    break

        # ── CAMP CONFIG SUMMARY ───────────────────────────────────────────────────────────
        st.markdown("---")
        st.subheader("📊 Camp Configuration")
        for _wk in [1, 2]:
            _wdf = df_y9_act[df_y9_act['Week'] == _wk]
            if _wdf.empty: continue
            _config = all_week_configs.get(_wk, {})
            _week_groups = all_week_results.get(_wk, {})
            _wl = "Week 1: Unwin & Hodgkin (12–18 Nov)" if _wk == 1 else "Week 2: Mather & Ransome (23–29 Nov)"
            _nc = sum(_config.values())
            with st.expander(f"📅 {_wl} — {len(_wdf)} students, {_nc} camps {'(6-camp plan)' if _nc == 6 else '(7-camp plan)'}", expanded=True):
                _cfg_rows = []
                for _ck, _ni in _config.items():
                    _mp = Y9_CAMP_DEFS[_ck]['max_per']
                    _grps = _week_groups.get(_ck, {})
                    _sz_a = len(_grps.get('A', []))
                    _sz_b = len(_grps.get('B', []))
                    if _ni == 1:
                        _cfg_rows.append({
                            'Camp': Y9_CAMP_DEFS[_ck]['label'], 'Instances': 1,
                            'Max Per Group': _mp, 'Assigned': _sz_a, 'Group B': '—'})
                    else:
                        _cfg_rows.append({
                            'Camp': Y9_CAMP_DEFS[_ck]['label'], 'Instances': 2,
                            'Max Per Group': _mp, 'Group A': _sz_a, 'Group B': _sz_b})
                st.dataframe(pd.DataFrame(_cfg_rows), hide_index=True, use_container_width=True)

        # ── RESULTS TABS (one per camp type) ──────────────────────────────────────────────
        st.markdown("---")
        st.subheader("🏕️ Draft Groups")

        _all_active = set()
        for _wg in all_week_results.values(): _all_active.update(_wg.keys())
        _ordered_camps = [k for k in Y9_CAMP_DEFS if k in _all_active]

        if not _ordered_camps:
            st.info("No groups generated yet. Upload both CSV files to begin.")
        else:
            _tabs = st.tabs([Y9_CAMP_DEFS[k]['label'] for k in _ordered_camps])
            _export_sheets = {}   # {sheet_name: df}

            for _tab, _ck in zip(_tabs, _ordered_camps):
                with _tab:
                    _mp = Y9_CAMP_DEFS[_ck]['max_per']

                    for _wk in [1, 2]:
                        _wdf = df_y9_act[df_y9_act['Week'] == _wk]
                        if _wdf.empty: continue
                        _week_groups = all_week_results.get(_wk, {})

                        _wt = ("📅 **Week 1: Unwin & Hodgkin** (12–18 November)"
                               if _wk == 1 else
                               "📅 **Week 2: Mather & Ransome** (23–29 November)")
                        st.markdown(_wt)

                        if _ck not in _week_groups:
                            st.info(f"This camp is not scheduled for Week {_wk}.")
                            st.markdown("---"); continue

                        _subgroups = _week_groups[_ck]
                        _has_b = len(_subgroups.get('B', [])) > 0
                        _sg_keys = ['A', 'B'] if _has_b else ['A']

                        _pref_w  = dict(zip(_wdf['Official Name'], _wdf['Camp Prefs']))
                        _fr_w    = dict(zip(_wdf['Official Name'], _wdf['Friends Requested']))

                        _sg_containers = st.columns(2) if _has_b else [st.container()]

                        for _sg_i, _sg in enumerate(_sg_keys):
                            _sg_list = _subgroups.get(_sg, [])
                            _cont = _sg_containers[_sg_i] if _has_b else _sg_containers[0]

                            with _cont:
                                _gt = (f"**Group {_sg}** — {len(_sg_list)}/{_mp} students"
                                       if _has_b else
                                       f"**{Y9_CAMP_DEFS[_ck]['label']}** — {len(_sg_list)}/{_mp} students")
                                if len(_sg_list) > _mp:
                                    st.markdown(f"🔴 {_gt} *(OVER CAPACITY)*")
                                else:
                                    st.markdown(_gt)

                                _rows = []
                                for _stud in _sg_list:
                                    _sr = _wdf[_wdf['Official Name'] == _stud]
                                    if _sr.empty: continue
                                    _sr = _sr.iloc[0]
                                    _resp = _sr['Responded']
                                    _friends_raw = _fr_w.get(_stud, [])
                                    _pref_rank = _pref_w.get(_stud, {}).get(_ck, '—')
                                    _friend_str = ', '.join(_friends_raw) if _friends_raw else '—'

                                    _rows.append({
                                        'Student ID':         _sr.get('Student ID', 'N/A'),
                                        'Student':            _stud,
                                        'Responded':          'Yes' if _resp else 'No',
                                        'Gender':             str(_sr.get('Gender', 'o')).upper(),
                                        'House':              str(_sr.get('House', '—')).title(),
                                        'Friend Requested':   _friend_str,
                                        'Pref Rank':          _pref_rank,
                                        'Phys Challenge':     _sr.get(_challenge_col, 'N/A') if _challenge_col and _resp else 'N/A',
                                        'Camping Skill':      _sr.get(_camping_col,   'N/A') if _camping_col   and _resp else 'N/A',
                                        'Overnight Hike':     _sr.get(_overnight_col, 'N/A') if _overnight_col and _resp else 'N/A',
                                        'Hardship Response':  _sr.get(_hardship_col,  'N/A') if _hardship_col  and _resp else 'N/A',
                                        'Swimming':           _sr.get(_swim_col,       'N/A') if _swim_col      and _resp else 'N/A',
                                        'White Water':        _sr.get(_ww_col,         'N/A') if _ww_col        and _resp else 'N/A',
                                        'Group Teamwork':     _sr.get(_teamwork_col,   'N/A') if _teamwork_col  and _resp else 'N/A',
                                    })

                                if _rows:
                                    _df_sg = pd.DataFrame(_rows)

                                    def _hl_y9(row, df_ref, wk, ck, find_name, fl, include_drafts=False):
                                        colors = [''] * len(row)
                                        if row['Student'] == find_name:
                                            return ['background-color: #85e085; font-weight: bold; color: black'] * len(row)
                                        if row['Responded'] == 'No':
                                            if include_drafts:
                                                # Draft student — full grey row
                                                return ['background-color: #d9d9d9; color: #555; font-style: italic'] * len(row)
                                            else:
                                                colors[df_ref.columns.get_loc('Responded')] = 'background-color: #ffffcc'

                                        # Pref rank colouring
                                        _pr_idx = df_ref.columns.get_loc('Pref Rank')
                                        try:
                                            _pr = int(row['Pref Rank'])
                                            if _pr == 1:   colors[_pr_idx] = 'background-color: #c8f7c5; font-weight: bold'
                                            elif _pr == 2: colors[_pr_idx] = 'background-color: #e8f8c5'
                                            elif _pr == 4: colors[_pr_idx] = 'background-color: #ffd9a0'
                                            elif _pr == 5: colors[_pr_idx] = 'background-color: #ffb3b3; font-weight: bold'
                                        except: pass

                                        # Friend colouring
                                        _f_idx = df_ref.columns.get_loc('Friend Requested')
                                        _s_idx = df_ref.columns.get_loc('Student')
                                        _f_str = str(row.get('Friend Requested', ''))
                                        if _f_str and _f_str != '—':
                                            _fl_list = [f.strip() for f in _f_str.split(',')]
                                            _all_together = all(
                                                fl.get(f, (None, None, None))[:2] == (wk, ck)
                                                for f in _fl_list)
                                            if not _all_together:
                                                colors[_f_idx] = 'background-color: #ffcccc'
                                                colors[_s_idx] = 'background-color: #ff9900; color: black; font-weight: bold'

                                        # Separation violation
                                        for _a, _b in st.session_state.y9_sep:
                                            if row['Student'] in (_a, _b):
                                                _other = _b if row['Student'] == _a else _a
                                                if fl.get(_other, (None, None, None))[:2] == (wk, ck):
                                                    colors[_s_idx] = 'background-color: #ff4d4d; color: white; font-weight: bold'

                                        # Skill columns: colour numeric values
                                        for _skill_col in ['Camping Skill', 'Overnight Hike', 'Hardship Response', 'Swimming', 'White Water', 'Group Teamwork', 'Phys Challenge']:
                                            if _skill_col in df_ref.columns:
                                                _sc_idx = df_ref.columns.get_loc(_skill_col)
                                                _val = str(row.get(_skill_col, '')).strip()
                                                _m = re.search(r'^\s*([1-9][0-9]?)(?:\D|$)', _val)
                                                if _m:
                                                    _num = int(_m.group(1))
                                                    if 1 <= _num <= 3:   colors[_sc_idx] = 'background-color: #a83232; color: white; font-weight: bold'
                                                    elif 4 <= _num <= 6: colors[_sc_idx] = 'background-color: #ff9933; color: black; font-weight: bold'
                                                    elif _num >= 7:      colors[_sc_idx] = 'background-color: #c8f7c5; color: black'

                                        return colors

                                    _styled = _df_sg.style.apply(
                                        _hl_y9, df_ref=_df_sg, wk=_wk, ck=_ck,
                                        find_name=y9_find, fl=full_lookup,
                                        include_drafts=y9_include_drafts, axis=1)
                                    st.dataframe(_styled, hide_index=True, use_container_width=True)

                                    _sheet_nm = f"{_ck}{'_'+_sg if _has_b else ''}_W{_wk}"[:31]
                                    _export_sheets[_sheet_nm] = _df_sg
                                else:
                                    st.info("No students in this group.")

                        st.markdown("---")

            # ── DASHBOARDS ────────────────────────────────────────────────────────────────
            st.markdown("---")
            _dcol1, _dcol2, _dcol3 = st.columns(3)
            with _dcol1:
                st.write("### ⚠️ At Risk of Isolation")
                st.caption("Students whose requested friend ended up in a different camp.")
                if isolated_y9:
                    st.dataframe(pd.DataFrame(isolated_y9), hide_index=True, use_container_width=True)
                else:
                    st.success("✅ All students with a friend request are together!")

            with _dcol2:
                st.write("### ❌ Missing Responses")
                # Use the full active pool (df_y9_act) so missing responses always appear
                # regardless of draft-mode setting
                _miss_y9 = df_y9_act[~df_y9_act['Responded']][
                    ['Student ID', 'Official Name', 'Email', 'House_stud']].copy()
                _miss_y9.rename(columns={'Official Name': 'Student', 'House_stud': 'House'}, inplace=True)
                _miss_y9['House'] = _miss_y9['House'].str.title()
                if not _miss_y9.empty:
                    st.dataframe(_miss_y9[['Student ID', 'Student', 'House']], hide_index=True, use_container_width=True)
                    st.write(f"**{len(_miss_y9)} student(s) yet to respond**")
                    _email_list = ', '.join(
                        e for e in _miss_y9['Email'].dropna().tolist()
                        if e and e.lower() not in ('', 'nan', 'none'))
                    if _email_list:
                        st.write("**📧 Email list (copy & paste into BCC):**")
                        st.code(_email_list, language=None)
                else:
                    st.success("✅ Everyone has responded!")

            with _dcol3:
                st.write("### 🚫 Not Attending")
                st.caption("Students marked as not attending this camp.")
                if _y9_na:
                    _na_display = df_y9[df_y9['Official Name'].isin(_y9_na)][
                        ['Student ID', 'Official Name', 'House_stud']].copy()
                    _na_display.rename(columns={'Official Name': 'Student', 'House_stud': 'House'}, inplace=True)
                    _na_display['House'] = _na_display['House'].str.title()
                    st.dataframe(_na_display, hide_index=True, use_container_width=True)
                    st.write(f"**{len(_y9_na)} student(s) not attending**")
                else:
                    st.info("No students marked as not attending.")
            if _export_sheets:
                st.markdown("---")
                _y9_out = io.BytesIO()
                with pd.ExcelWriter(_y9_out, engine='openpyxl') as _writer:
                    # Ordered sheets: by camp type, then week, then subgroup
                    _ordered_sheets = sorted(
                        _export_sheets.keys(),
                        key=lambda nm: (
                            list(Y9_CAMP_DEFS.keys()).index(nm.split('_')[0])
                            if nm.split('_')[0] in Y9_CAMP_DEFS else 99,
                            nm
                        ))
                    for _snm in _ordered_sheets:
                        _df_sh = _export_sheets[_snm]
                        _df_sh.to_excel(_writer, sheet_name=_snm, index=False)
                        _ws = _writer.sheets[_snm]
                        for _col in _ws.columns:
                            _ws.column_dimensions[_col[0].column_letter].width = min(
                                max((len(str(cell.value)) for cell in _col), default=0) + 2, 45)

                    # Isolation report sheet
                    if isolated_y9:
                        pd.DataFrame(isolated_y9).to_excel(_writer, sheet_name="Isolation Report", index=False)
                        _ws_iso = _writer.sheets["Isolation Report"]
                        for _col in _ws_iso.columns:
                            _ws_iso.column_dimensions[_col[0].column_letter].width = max(
                                (len(str(cell.value)) for cell in _col), default=0) + 2

                    # Not Attending sheet
                    if _y9_na:
                        _na_df = df_y9[df_y9['Official Name'].isin(_y9_na)][
                            ['Student ID', 'Official Name', 'House_stud', 'Responded']].copy()
                        _na_df.rename(columns={'Official Name': 'Student', 'House_stud': 'House'}, inplace=True)
                        _na_df.to_excel(_writer, sheet_name="Not Attending", index=False)
                        _ws_na = _writer.sheets["Not Attending"]
                        for _col in _ws_na.columns:
                            _ws_na.column_dimensions[_col[0].column_letter].width = max(
                                (len(str(cell.value)) for cell in _col), default=0) + 2

                _y9_out.seek(0)
                st.download_button(
                    label="📥 Download Y9 Journey Groups (Excel)",
                    data=_y9_out,
                    file_name="Y9_Journey_Groups.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )