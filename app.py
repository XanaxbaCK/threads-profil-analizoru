import json
import re
import io
from pathlib import Path
from typing import Any, List, Dict
from datetime import datetime
import streamlit as st
import xlsxwriter

# --- MOBİL VE GENİŞ EKRAN UYUMLULUK AYARI ---
st.set_page_config(
    page_title="Threads Profil Analizörü",
    page_icon="🎯",
    layout="centered",
    initial_sidebar_state="collapsed"
)
# --- ÇOKLU DİL SÖZLÜĞÜ (TR / EN / DE) ---
DIL_PAKETI = {
    "TR": {
        "main_title": "THREADS TÜRKİYE",
        "main_sub": "YEREL VE GÜVENLİ ÇİFT YÖNLÜ PROFİL TAKİPÇİ SİSTEMİ",
        "main_hashtag": "#bendeğilbizyaptık",
        "load_following": "Takip Ettiklerinizi Yükleyin (following.json)",
        "load_followers": "Takipçilerinizi Yükleyin (followers.json)",
        "btn_analyze": "ANALİZİ BAŞLAT",
        "tab_unfollowers": "Beni Takip Etmeyenler",
        "tab_fans": "Geri Takip Etmediklerim",
        "tab_ghosts": "Hayalet (Ghost) Hesaplar",
        "input_error_msg": "Analiz için gerekli iki kaynak dosya da yüklenmelidir.",
        "parse_error_msg": "Yüklenen JSON şeması motor tarafından çözümlenemedi.",
        "success_msg": "Dosyalar tarandı, Excel ve TXT raporları köprü linkleriyle üretildi!",
        "perfect_sync": "🎉 [KUSURSUZ SENKRONİZASYON]: Herkes sizi geri takip ediyor!",
        "no_fans": "🎯 [HAYRAN YOK]: Takip ettiğiniz herkesi siz de geri takip ediyorsunuz.",
        "no_ghosts": "🛡️ [TEMİZ PROFİL]: Profilinizde hayalet veya bot hesap algılanmadı.",
        "download_excel": "📥 Excel Analiz Raporunu İndir",
        "summary_title": "📊 PROFİL SAĞLIK ÖZETİ",
        "health_score": "Sağlık Skoru"
    },
    "EN": {
        "main_title": "THREADS GLOBAL",
        "main_sub": "LOCAL AND SECURE BI-DIRECTIONAL PROFILE ANALYSIS SYSTEM",
        "main_hashtag": "#notmebutwe",
        "load_following": "Load Those You Follow (following.json)",
        "load_followers": "Load Your Followers (followers.json)",
        "btn_analyze": "START ANALYSIS",
        "tab_unfollowers": "Not Following Me Back",
        "tab_fans": "I Am Not Following Back",
        "tab_ghosts": "Ghost / Inactive Accounts",
        "input_error_msg": "Both required source files must be uploaded for analysis.",
        "parse_error_msg": "The uploaded JSON schema could not be resolved by the engine.",
        "success_msg": "Files scanned, Excel and TXT reports generated with clickable links!",
        "perfect_sync": "🎉 [PERFECT SYNC]: Everyone is following you back!",
        "no_fans": "🎯 [NO FANS]: You are following back everyone who follows you.",
        "no_ghosts": "🛡️ [CLEAN PROFILE]: No ghost or bot accounts detected on your profile.",
        "download_excel": "📥 Download Excel Analysis Report",
        "summary_title": "📊 PROFILE HEALTH SUMMARY",
        "health_score": "Health Score"
    },
    "DE": {
        "main_title": "THREADS GLOBAL",
        "main_sub": "LOKALES UND SICHERES BIDIREKTIONALES PROFIL-ANALYSESYSTEM",
        "main_hashtag": "#nichtichsondernwir",
        "load_following": "Laden Sie die, denen Sie folgen (following.json)",
        "load_followers": "Laden Sie Ihre Follower (followers.json)",
        "btn_analyze": "ANALYSE STARTEN",
        "tab_unfollowers": "Folgen mir nicht zurück",
        "tab_fans": "Ich folge nicht zurück",
        "tab_ghosts": "Geister / Inaktive Konten",
        "input_error_msg": "Für die Analyse müssen beide Quelldateien hochgeladen werden.",
        "parse_error_msg": "Das hochgeladene JSON-Schema konnte nicht aufgelöst werden.",
        "success_msg": "Dateien erfolgreich gescannt, Berichte als Excel und TXT exportiert!",
        "perfect_sync": "🎉 [PERFEKTE SYNCHRONISATION]: Jeder folgt Ihnen zurück!",
        "no_fans": "🎯 [KEINE FANS]: Sie folgen jedem zurück, der Ihnen folgt.",
        "no_ghosts": "🛡️ [SAUBERES PROFIL]: Keine Geister- oder Bot-Konten auf Ihrem Profil erkannt.",
        "download_excel": "📥 Excel-Analysebericht herunterladen",
        "summary_title": "📊 PROFIL-GESUNDHEITSÜBERSICHT",
        "health_score": "Gesundheitsscore"
    }
}
class AnalizMotoru:
    """Kullanıcı adlarını, zaman damgalarını, bot riskini ve Sağlık Skorunu hesaplayan ana motor."""
    @staticmethod
    def akilli_süre_ayristir(data: Any) -> Dict[str, int]:
        results: Dict[str, int] = {}
        if isinstance(data, dict):
            if "string_list_data" in data and isinstance(data["string_list_data"], list):
                for item in data["string_list_data"]:
                    if isinstance(item, dict) and "value" in item:
                        val = str(item["value"]).strip()
                        if val and not val.isdigit() and not val.startswith("http"):
                            results[val] = item.get("timestamp", 0)
                return results
        if isinstance(data, dict):
            if "value" in data and isinstance(data["value"], str):
                val = data["value"].strip()
                if val and not val.isdigit() and not val.startswith(("http", "202")):
                    results[val] = data.get("timestamp", 0)
            for value in data.values():
                results.update(AnalizMotoru.akilli_süre_ayristir(value))
        elif isinstance(data, list):
            for item in data:
                results.update(AnalizMotoru.akilli_süre_ayristir(item))
        return results

    @staticmethod
    def zaman_metnine_cevir(timestamp: int) -> str:
        if timestamp == 0: return "Bilinmiyor"
        try:
            fark = datetime.now() - datetime.fromtimestamp(timestamp)
            gun = fark.days
            if gun < 30: return f"{gun} Gün"
            ay = gun // 30
            if ay < 12: return f"{ay} Ay"
            return f"{ay // 12} Yıl {ay % 12} Ay"
        except Exception: return "Bilinmiyor"

    @staticmethod
    def bot_ve_pasiflik_kontrolü(username: str, timestamp: int) -> bool:
        """Kullanıcı adındaki ardışık sayıları ve pasiflik sürelerini ölçen risk dedektörü."""
        if not username: return False
        if re.search(r'\d{3,}', username):
            return True
        if timestamp > 0:
            fark_gun = (datetime.now() - datetime.fromtimestamp(timestamp)).days
            if fark_gun > 540:
                return True
        return False
# --- MOBİL ARAYÜZ YAPILANDIRMASI VE DİL SEÇİMİ ---
st.title("🎯 Threads Profil Analizörü")

col_lang, col_hashtag = st.columns([1, 2])
with col_lang:
    aktif_dil = st.selectbox("🌐 Language / Dil", ["TR", "EN", "DE"])

with col_hashtag:
    st.markdown(f"<h4 style='text-align: right; color: #3a7ebf; margin-top: 5px;'>{DIL_PAKETI[aktif_dil]['main_hashtag']}</h4>", unsafe_allow_html=True)

st.markdown(f"### {DIL_PAKETI[aktif_dil]['main_title']}")
st.caption(DIL_PAKETI[aktif_dil]['main_sub'])
st.divider()

# --- MOBİL UYUMLU DOSYA YÜKLEME ALANLARI ---
uploaded_following = st.file_uploader(DIL_PAKETI[aktif_dil]['load_following'], type=["json"])
uploaded_followers = st.file_uploader(DIL_PAKETI[aktif_dil]['load_followers'], type=["json"], accept_multiple_files=True)

btn_trigger = st.button(DIL_PAKETI[aktif_dil]['btn_analyze'], use_container_width=True, type="primary")
if btn_trigger:
    if not uploaded_following or not uploaded_followers:
        st.warning(DIL_PAKETI[aktif_dil]['input_error_msg'])
    else:
        try:
            # Belleğe yüklenen dosyaları okuma
            following_raw = json.loads(uploaded_following.read().decode("utf-8"))
            global_following_map = AnalizMotoru.akilli_süre_ayristir(following_raw)
            
            global_followers_map = {}
            for u_file in uploaded_followers:
                followers_raw = json.loads(u_file.read().decode("utf-8"))
                global_followers_map.update(AnalizMotoru.akilli_süre_ayristir(followers_raw))
                
            following_set = set(global_following_map.keys())
            followers_set = set(global_followers_map.keys())
            
            if not following_set or not followers_set:
                st.error(DIL_PAKETI[aktif_dil]['parse_error_msg'])
            else:
                unfollowers = following_set - followers_set
                fans = followers_set - following_set
                ghosts = {u for u, ts in global_followers_map.items() if AnalizMotoru.bot_ve_pasiflik_kontrolü(u, ts)}
                # Sağlık Skoru Hesaplama Mantığı
                toplam_bağ = len(following_set) + len(followers_set)
                if toplam_bağ > 0:
                    ceza_puanı = (len(unfollowers) / len(following_set)) * 40 if len(following_set) > 0 else 0
                    ghost_ceza = (len(ghosts) / len(followers_set)) * 20 if len(followers_set) > 0 else 0
                    denge_puanı = (min(len(followers_set), len(following_set)) / max(len(followers_set), len(following_set))) * 40
                    health_score = max(0, min(100, int(100 - ceza_puanı - ghost_ceza + (denge_puanı * 0.1))))
                else:
                    health_score = 100
                
                durum_str = "MÜKEMMEL" if health_score > 85 else "STABİL" if health_score > 60 else "RİSKLİ"
                
                # --- MOBİL PANEL ÖZET KARTLARI ---
                st.success(DIL_PAKETI[aktif_dil]['success_msg'])
                st.subheader(DIL_PAKETI[aktif_dil]['summary_title'])
                
                m1, m2, m3 = st.columns(3)
                m1.metric(DIL_PAKETI[aktif_dil]['health_score'], f"%{health_score}", durum_str)
                m2.metric("Following", len(following_set))
                m3.metric("Followers", len(followers_set))
                # --- BELLEKTE EXCEL OLUŞTURMA MOTORU (XLSXWRITER) ---
                output_excel = io.BytesIO()
                workbook = xlsxwriter.Workbook(output_excel)
                
                header_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#1F4E78', 'border': 1, 'align': 'center'})
                link_format = workbook.add_format({'font_color': 'blue', 'underline': True})
                text_format = workbook.add_format({'align': 'left'})
                
                # Sekme Metin Havuzları
                unf_list_str = []
                fans_list_str = []
                ghosts_list_str = []
                # 1. Sayfa: Beni Takip Etmeyenler
                sheet_unf = workbook.add_worksheet("Beni Takip Etmeyenler")
                sheet_unf.write_row('A1', ['No', 'Kullanıcı Adı', 'Profil Linki', 'Süre'], header_format)
                sorted_unf = sorted(unfollowers, key=lambda u: global_following_map.get(u, 0))
                for idx, user in enumerate(sorted_unf, 1):
                    süre = AnalizMotoru.zaman_metnine_cevir(global_following_map.get(user, 0))
                    p_url = f"https://threads.com/@{user}"
                    unf_list_str.append(f"[{idx:03d}] @{user:<20} ⌛ {süre}")
                    sheet_unf.write(idx, 0, idx)
                    sheet_unf.write(idx, 1, f"@{user}", text_format)
                    sheet_unf.write_url(idx, 2, p_url, link_format, string=p_url)
                    sheet_unf.write(idx, 3, süre, text_format)
                sheet_unf.set_column('A:A', 5); sheet_unf.set_column('B:B', 20); sheet_unf.set_column('C:C', 45); sheet_unf.set_column('D:D', 20)

                # 2. Sayfa: Geri Takip Etmediklerim
                sheet_fans = workbook.add_worksheet("Geri Takip Etmediklerim")
                sheet_fans.write_row('A1', ['No', 'Kullanıcı Adı', 'Profil Linki', 'Süre'], header_format)
                sorted_fans = sorted(fans, key=lambda f: global_followers_map.get(f, 0))
                for idx, user in enumerate(sorted_fans, 1):
                    süre = AnalizMotoru.zaman_metnine_cevir(global_followers_map.get(user, 0))
                    p_url = f"https://threads.com/@{user}"
                    fans_list_str.append(f"[{idx:03d}] @{user:<20} ⌛ {süre}")
                    sheet_fans.write(idx, 0, idx)
                    sheet_fans.write(idx, 1, f"@{user}", text_format)
                    sheet_fans.write_url(idx, 2, p_url, link_format, string=p_url)
                    sheet_fans.write(idx, 3, süre, text_format)
                sheet_fans.set_column('A:A', 5); sheet_fans.set_column('B:B', 20); sheet_fans.set_column('C:C', 45); sheet_fans.set_column('D:D', 20)
                # 3. Sayfa: Hayalet Hesaplar
                sheet_gh = workbook.add_worksheet("Hayalet Hesaplar")
                sheet_gh.write_row('A1', ['No', 'Kullanıcı Adı', 'Profil Linki', 'Süre'], header_format)
                sorted_gh = sorted(ghosts, key=lambda g: global_followers_map.get(g, 0))
                for idx, user in enumerate(sorted_gh, 1):
                    süre = AnalizMotoru.zaman_metnine_cevir(global_followers_map.get(user, 0))
                    p_url = f"https://threads.com/@{user}"
                    ghosts_list_str.append(f"[{idx:03d}] @{user:<20} ⌛ {süre}")
                    sheet_gh.write(idx, 0, idx)
                    sheet_gh.write(idx, 1, f"@{user}", text_format)
                    sheet_gh.write_url(idx, 2, p_url, link_format, string=p_url)
                    sheet_gh.write(idx, 3, süre, text_format)
                sheet_gh.set_column('A:A', 5); sheet_gh.set_column('B:B', 20); sheet_gh.set_column('C:C', 45); sheet_gh.set_column('D:D', 20)
                
                workbook.close()
                output_excel.seek(0)
                
                # --- EXCEL İNDİRME BUTONU ---
                st.download_button(
                    label=DIL_PAKETI[aktif_dil]['download_excel'],
                    data=output_excel,
                    file_name="Threads_Detayli_Analiz_Raporu.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                               # --- WEB SEKME GÖRÜNÜMLERİ ---
                t1, t2, t3 = st.tabs([DIL_PAKETI[aktif_dil]['tab_unfollowers'], DIL_PAKETI[aktif_dil]['tab_fans'], DIL_PAKETI[aktif_dil]['tab_ghosts']])
                
                with t1:
                    if unfollowers:
                        for index, user in enumerate(sorted_unf, 1):
                            süre = AnalizMotoru.zaman_metnine_cevir(global_following_map.get(user, 0))
                            p_url = f"https://threads.com/@{user}"
                            # Her satırı tıklanabilir mavi bir link haline getiriyoruz
                            st.markdown(f"[{index:03d}] 🔗 [@{user}]({p_url}) &nbsp;&nbsp;&nbsp;&nbsp; ⌛ {süre}")
                    else:
                        st.info(DIL_PAKETI[aktif_dil]['perfect_sync'])
                        
                with t2:
                    if fans:
                        for index, user in enumerate(sorted_fans, 1):
                            süre = AnalizMotoru.zaman_metnine_cevir(global_followers_map.get(user, 0))
                            p_url = f"https://threads.com/@{user}"
                            st.markdown(f"[{index:03d}] 🔗 [@{user}]({p_url}) &nbsp;&nbsp;&nbsp;&nbsp; ⌛ {süre}")
                    else:
                        st.info(DIL_PAKETI[aktif_dil]['no_fans'])
                        
                with t3:
                    if ghosts:
                        for index, user in enumerate(sorted_gh, 1):
                            süre = AnalizMotoru.zaman_metnine_cevir(global_followers_map.get(user, 0))
                            p_url = f"https://threads.com/@{user}"
                            st.markdown(f"[{index:03d}] 🔗 [@{user}]({p_url}) &nbsp;&nbsp;&nbsp;&nbsp; ⌛ {süre}")
                    else:
                        st.info(DIL_PAKETI[aktif_dil]['no_ghosts'])
                        
        except Exception as e:
            st.error(f"Sistem Hatası: {str(e)}")

if __name__ == "__main__":
    # Streamlit uygulamalarında ana döngü otomatik tetiklendiği için bu kısım boş kalabilir veya kaldırılabilir.
    pass

