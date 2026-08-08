import json
import re
import io
import urllib.parse
from pathlib import Path
from typing import Any, List, Dict
from datetime import datetime
import streamlit as st
import xlsxwriter

# --- PREMIUM SİBER GECE MODU VE GENİŞ EKRAN AYARI ---
st.set_page_config(
    page_title="Threads Cyber Tracking System",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# --- SİBER NEON CSS ENJEKSİYONU ---
st.markdown("""
    <style>
    .stApp {
        background-color: #030712 !important;
    }
    h1, h2, h3, h4, h5, h6, label, p, span {
        color: #00ff66 !important;
        font-family: 'Consolas', monospace !important;
        text-shadow: 0 0 5px #00ff66, 0 0 10px #00ff66 !important;
    }
    .stButton>button {
        background-color: #111827 !important;
        color: #00ff66 !important;
        border: 2px solid #00ff66 !important;
        box-shadow: 0 0 10px #00ff66 !important;
        font-family: 'Consolas', monospace !important;
        font-weight: bold !important;
    }
    .stButton>button:hover {
        background-color: #00ff66 !important;
        color: #030712 !important;
        box-shadow: 0 0 20px #00ff66 !important;
    }
    .stTabs [data-baseweb="tab-list"] {
        background-color: #0b0f19 !important;
        border: 1px solid #1f2937 !important;
        border-radius: 8px !important;
    }
    .stTabs [data-baseweb="tab"] {
        color: #6b7280 !important;
    }
    .stTabs [aria-selected="true"] {
        color: #00ff66 !important;
        text-shadow: 0 0 5px #00ff66 !important;
    }
    </style>
    """, unsafe_allow_html=True)
# --- ÇOKLU DİL SÖZLÜĞÜ (TR / EN / DE) ---
DIL_PAKETI = {
    "TR": {
        "main_title": "⚡ THREADS CYBER // MURAT & ESRA (CAN&KAN)",
        "main_sub": "SYSTEM ENG: YEREL VE GÜVENLİ ÇİFT YÖNLÜ SİBER PROFİL ANALİZİ",
        "main_hashtag": "#bendeğilbizyaptık",
        "load_following": "Takip Ettiklerinizi Yükleyin (following.json)",
        "load_followers": "Takipçilerinizi Yükleyin (followers.json)",
        "btn_analyze": "⚡ SİBER TARAMAYI BAŞLAT ⚡",
        "tab_unfollowers": "[!] Beni Takip Etmeyenler",
        "tab_fans": "[+] Geri Takip Etmediklerim",
        "tab_ghosts": "[💀] Hayalet (Ghost) Hesaplar",
        "input_error_msg": "CRITICAL ERROR: Analiz için iki kaynak dosya da yüklenmelidir.",
        "parse_error_msg": "DECODE ERROR: Yüklenen JSON şeması motor tarafından çözümlenemedi.",
        "success_msg": "SUCCESS: Dosyalar tarandı, veri havuzu köprü linkleriyle üretildi!",
        "perfect_sync": "🎉 [SAFE LOG]: Herkes sizi geri takip ediyor, siber uyum kusursuz!",
        "no_fans": "🎯 [SAFE LOG]: Takip ettiğiniz herkesi siz de geri takip ediyorsunuz.",
        "no_ghosts": "🛡️ [SAFE LOG]: Profilinizde hayalet veya bot hesap algılanmadı.",
        "download_excel": "📥 Cyber Excel Analiz Raporunu İndir",
        "summary_title": "📊 SİBER PROFİL SAĞLIK ÖZETİ",
        "health_score": "Siber Sağlık Skoru",
        "guide_title": "📖 Threads Verileri Nasıl İndirilir? (Siber Kılavuz)",
        "guide_step1": "1️⃣ **Instagram/Threads** uygulamasını açın ve **Ayarlar -> Hesaplar Merkezi** bölümüne girin.",
        "guide_step2": "2️⃣ **Bilgilerin ve İzinlerin -> Bilgilerini İndir** adımlarını takip edin.",
        "guide_step3": "3️⃣ **Indirme Talep Et** butonuna basın ve sadece **Threads** seçeneğini işaretleyin.",
        "guide_step4": "4️⃣ Dosya formatını **JSON** (ÖNEMLİ!), medya kalitesini **Düşük** seçip talebi onaylayın.",
        "guide_step5": "5️⃣ E-postanıza gelen `.zip` dosyasını indirin veConnections klasörüne çıkartın.",
        "guide_step6": "6️⃣ Klasörün içindeki `connections/followers_and_following` yoluna giderek **`followers.json`** ve **`following.json`** dosyalarını yükleyin.",
        "contact_btn": "💬 SİBER YAPIMCI İLE İLETİŞİME GEÇ (@muratsenr)",
        "player_title": "🎵 Arka Plan Müziği: Cankan - Yaranamadım (Siber Versiyon)"
    },
    "EN": {
        "main_title": "⚡ THREADS CYBER GLOBAL",
        "main_sub": "SYSTEM ENG: LOCAL & SECURE BI-DIRECTIONAL CYBER ANALYSIS",
        "main_hashtag": "#notmebutwe",
        "load_following": "Load Those You Follow (following.json)",
        "load_followers": "Load Your Followers (followers.json)",
        "btn_analyze": "⚡ LAUNCH CYBER SCAN ⚡",
        "tab_unfollowers": "[!] Not Following Me Back",
        "tab_fans": "[+] I Am Not Following Back",
        "tab_ghosts": "[💀] Ghost / Inactive Accounts",
        "input_error_msg": "CRITICAL ERROR: Both required source files must be uploaded.",
        "parse_error_msg": "DECODE ERROR: The uploaded JSON schema could not be resolved.",
        "success_msg": "SUCCESS: Files scanned, cyber data generated with links!",
        "perfect_sync": "🎉 [SAFE LOG]: Everyone is following you back!",
        "no_fans": "🎯 [SAFE LOG]: You are following back everyone who follows you.",
        "no_ghosts": "🛡️ [SAFE LOG]: No ghost or bot accounts detected.",
        "download_excel": "📥 Download Cyber Excel Report",
        "summary_title": "📊 CYBER PROFILE SUMMARY",
        "health_score": "Cyber Health Score",
        "guide_title": "📖 How to Download Threads Data? (Cyber Guide)",
        "guide_step1": "1️⃣ Open **Instagram/Threads**, go to **Settings -> Accounts Center**.",
        "guide_step2": "2️⃣ Follow **Your Information and Permissions -> Download Your Information**.",
        "guide_step3": "3️⃣ Click **Request a Download** and select only **Threads**.",
        "guide_step4": "4️⃣ Choose format as **JSON**, media quality as **Low** and submit.",
        "guide_step5": "5️⃣ Download the `.zip` file from your email and extract it.",
        "guide_step6": "6️⃣ Go to `connections/followers_and_following` folder and upload files below.",
        "contact_btn": "💬 CONTACT CYBER DEVELOPER (@muratsenr)",
        "player_title": "🎵 Background Music: Cankan - Yaranamadım"
    },
    "DE": {
        "main_title": "⚡ THREADS CYBER GLOBAL",
        "main_sub": "SYSTEM ENG: LOKALE & SICHERE BIDIREKTIONALE CYBER-ANALYSE",
        "main_hashtag": "#nichtichsondernwir",
        "load_following": "Laden Sie die, denen Sie folgen (following.json)",
        "load_followers": "Laden Sie Ihre Follower (followers.json)",
        "btn_analyze": "⚡ CYBER-SCAN STARTEN ⚡",
        "tab_unfollowers": "[!] Folgen mir nicht zurück",
        "tab_fans": "[+] Ich folge nicht zurück",
        "tab_ghosts": "[💀] Geister / Inaktive Konten",
        "input_error_msg": "CRITICAL ERROR: Für die Analyse müssen beide Quelldateien hochgeladen werden.",
        "parse_error_msg": "DECODE ERROR: Das hochgeladene JSON-Schema konnte nicht aufgelöst werden.",
        "success_msg": "SUCCESS: Dateien erfolgreich gescannt, Berichte exportiert!",
        "perfect_sync": "🎉 [SAFE LOG]: Jeder folgt Ihnen zurück!",
        "no_fans": "🎯 [SAFE LOG]: Sie folgen jedem zurück, der Ihnen folgt.",
        "no_ghosts": "🛡️ [SAFE LOG]: Keine Geister- oder Bot-Konten auf Ihrem Profil.",
        "download_excel": "📥 Cyber-Excel-Analysebericht herunterladen",
        "summary_title": "📊 CYBER-PROFILÜBERSICHT",
        "health_score": "Cyber-Gesundheitsscore",
        "guide_title": "📖 Wie lade ich Threads-Daten herunter? (Cyber-Handbuch)",
        "guide_step1": "1️⃣ Öffnen Sie **Instagram/Threads**, gehen Sie zu **Einstellungen -> Kontenübersicht**.",
        "guide_step2": "2️⃣ Folgen Sie **Deine Informationen und Berechtigungen -> Deine Informationen herunterladen**.",
        "guide_step3": "3️⃣ Klicken Sie auf **Download anfordern** und wählen Sie nur **Threads** aus.",
        "guide_step4": "4️⃣ Wählen Sie das Format **JSON** und die Medienqualität **Niedrig**.",
        "guide_step5": "5️⃣ Laden Sie die `.zip`-Datei herunter und entpacken Sie sie.",
        "guide_step6": "6️⃣ Gehen Sie zum Ordner `connections/followers_and_following` und laden Sie Dateien hoch.",
        "contact_btn": "💬 CYBER-ENTWICKLER KONTAKTIEREN (@muratsenr)",
        "player_title": "🎵 Hintergrundmusik: Cankan - Yaranamadım"
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
st.markdown("<h1>⚡ Threads Cyber Tracking System</h1>", unsafe_allow_html=True)

col_lang, col_hashtag = st.columns(2)
with col_lang:
    aktif_dil = st.selectbox("🌐 Language / Dil", ["TR", "EN", "DE"])

with col_hashtag:
    st.markdown(f"<h4 style='text-align: right; color: #00ff66; margin-top: 5px;'>{DIL_PAKETI[aktif_dil]['main_hashtag']}</h4>", unsafe_allow_html=True)

st.markdown(f"### {DIL_PAKETI[aktif_dil]['main_title']}")
st.caption(DIL_PAKETI[aktif_dil]['main_sub'])
st.divider()

# --- 🎛️ GÜVENLİ MÜZİK BUTONU PANELİ ---
st.caption(DIL_PAKETI[aktif_dil]['player_title'])
st.link_button(
    label="▶️ YAPARKEN DİNLERSİNİZ YA (Göndermeli Şarkı)",
    url="https://youtube.com",
    use_container_width=True
)

# --- 📖 RESİMLİ / ADIM ADIM KULLANIM KILAVUZU PANELİ ---
with st.expander(DIL_PAKETI[aktif_dil]['guide_title'], expanded=False):
    st.markdown(DIL_PAKETI[aktif_dil]['guide_step1'])
    st.markdown(DIL_PAKETI[aktif_dil]['guide_step2'])
    st.markdown(DIL_PAKETI[aktif_dil]['guide_step3'])
    st.markdown(DIL_PAKETI[aktif_dil]['guide_step4'])
    st.markdown(DIL_PAKETI[aktif_dil]['guide_step5'])
    st.markdown(DIL_PAKETI[aktif_dil]['guide_step6'])
    st.info("ℹ️ SECURITY NOTICE: Raporlar yerel işlenir, verileriniz asla sunucuya kaydedilmez.")

st.write("") 

# --- MOBİL UYUMLU DOSYA YÜKLEME ALANLARI ---
uploaded_following = st.file_uploader(DIL_PAKETI[aktif_dil]['load_following'], type=["json"])
uploaded_followers = st.file_uploader(DIL_PAKETI[aktif_dil]['load_followers'], type=["json"], accept_multiple_files=True)

btn_trigger = st.button(DIL_PAKETI[aktif_dil]['btn_analyze'], use_container_width=True, type="primary")
if btn_trigger:
    if not uploaded_following or not uploaded_followers:
        st.warning(DIL_PAKETI[aktif_dil]['input_error_msg'])
    else:
        try:
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

                toplam_bağ = len(following_set) + len(followers_set)
                if toplam_bağ > 0:
                    ceza_puanı = (len(unfollowers) / len(following_set)) * 40 if len(following_set) > 0 else 0
                    ghost_ceza = (len(ghosts) / len(followers_set)) * 20 if len(followers_set) > 0 else 0
                    denge_puanı = (min(len(followers_set), len(following_set)) / max(len(followers_set), len(following_set))) * 40
                    health_score = max(0, min(100, int(100 - ceza_puanı - ghost_ceza + (denge_puanı * 0.1))))
                else:
                    health_score = 100
                
                durum_str = "MÜKEMMEL" if health_score > 85 else "STABİL" if health_score > 60 else "RİSKLİ"
                
                st.success(DIL_PAKETI[aktif_dil]['success_msg'])
                st.subheader(DIL_PAKETI[aktif_dil]['summary_title'])
                
                m1, m2, m3 = st.columns(3)
                m1.metric(DIL_PAKETI[aktif_dil]['health_score'], f"%{health_score}", durum_str)
                m2.metric("Following", len(following_set))
                m3.metric("Followers", len(followers_set))
                
                # --- BELLEKTE EXCEL OLUŞTURMA MOTORU (XLSXWRITER) ---
                output_excel = io.BytesIO()
                workbook = xlsxwriter.Workbook(output_excel)
                
                header_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#111827', 'border': 1, 'align': 'center'})
                link_format = workbook.add_format({'font_color': '#00ff66', 'underline': True})
                text_format = workbook.add_format({'align': 'left'})
                # 1. Sayfa: Beni Takip Etmeyenler
                sheet_unf = workbook.add_worksheet("Beni Takip Etmeyenler")
                sheet_unf.write_row('A1', ['No', 'Kullanıcı Adı', 'Profil Linki', 'Süre'], header_format)
                sorted_unf = sorted(unfollowers, key=lambda x: global_following_map.get(x, 0))
                for idx, user in enumerate(sorted_unf, 1):
                    süre = AnalizMotoru.zaman_metnine_cevir(global_following_map.get(user, 0))
                    p_url = f"https://threads.com@{user}"
                    sheet_unf.write(idx, 0, idx)
                    sheet_unf.write(idx, 1, f"@{user}", text_format)
                    sheet_unf.write_url(idx, 2, p_url, link_format, string=p_url)
                    sheet_unf.write(idx, 3, süre, text_format)
                sheet_unf.set_column('A:A', 5); sheet_unf.set_column('B:B', 20); sheet_unf.set_column('C:C', 45); sheet_unf.set_column('D:D', 20)

                # 2. Sayfa: Geri Takip Etmediklerim
                sheet_fans = workbook.add_worksheet("Geri Takip Etmediklerim")
                sheet_fans.write_row('A1', ['No', 'Kullanıcı Adı', 'Profil Linki', 'Süre'], header_format)
                sorted_fans = sorted(fans, key=lambda x: global_followers_map.get(x, 0))
                for idx, user in enumerate(sorted_fans, 1):
                    süre = AnalizMotoru.zaman_metnine_cevir(global_followers_map.get(user, 0))
                    p_url = f"https://threads.com@{user}"
                    sheet_fans.write(idx, 0, idx)
                    sheet_fans.write(idx, 1, f"@{user}", text_format)
                    sheet_fans.write_url(idx, 2, p_url, link_format, string=p_url)
                    sheet_fans.write(idx, 3, süre, text_format)
                sheet_fans.set_column('A:A', 5); sheet_fans.set_column('B:B', 20); sheet_fans.set_column('C:C', 45); sheet_fans.set_column('D:D', 20)

                # 3. Sayfa: Hayalet Hesaplar
                sheet_gh = workbook.add_worksheet("Hayalet Hesaplar")
                sheet_gh.write_row('A1', ['No', 'Kullanıcı Adı', 'Profil Linki', 'Süre'], header_format)
                sorted_gh = sorted(ghosts, key=lambda x: global_followers_map.get(x, 0))
                for idx, user in enumerate(sorted_gh, 1):
                    süre = AnalizMotoru.zaman_metnine_cevir(global_followers_map.get(user, 0))
                    p_url = f"https://threads.com@{user}"
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
                    file_name="Threads_Cyber_Raporu.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                # --- WEB SEKME GÖRÜNÜMLERİ (Tıklanabilir Siber Linkler) ---
                t1, t2, t3 = st.tabs([DIL_PAKETI[aktif_dil]['tab_unfollowers'], DIL_PAKETI[aktif_dil]['tab_fans'], DIL_PAKETI[aktif_dil]['tab_ghosts']])
                
                with t1:
                    if unfollowers:
                        for index, user in enumerate(sorted_unf, 1):
                            süre = AnalizMotoru.zaman_metnine_cevir(global_following_map.get(user, 0))
                            p_url = f"https://threads.com@{user}"
                            st.markdown(f"[{index:03d}] 🟢 [@{user}]({p_url}) &nbsp;&nbsp;&nbsp;&nbsp; ⌛ {süre}")
                    else:
                        st.info(DIL_PAKETI[aktif_dil]['perfect_sync'])
                        
                with t2:
                    if fans:
                        for index, user in enumerate(sorted_fans, 1):
                            süre = AnalizMotoru.zaman_metnine_cevir(global_followers_map.get(user, 0))
                            p_url = f"https://threads.com@{user}"
                            st.markdown(f"[{index:03d}] 🟢 [@{user}]({p_url}) &nbsp;&nbsp;&nbsp;&nbsp; ⌛ {süre}")
                    else:
                        st.info(DIL_PAKETI[aktif_dil]['no_fans'])
                        
                with t3:
                    if ghosts:
                        for index, user in enumerate(sorted_gh, 1):
                            süre = AnalizMotoru.zaman_metnine_cevir(global_followers_map.get(user, 0))
                            p_url = f"https://threads.com@{user}"
                            st.markdown(f"[{index:03d}] 🟢 [@{user}]({p_url}) &nbsp;&nbsp;&nbsp;&nbsp; ⌛ {süre}")
                    else:
                        st.info(DIL_PAKETI[aktif_dil]['no_ghosts'])
                        
        except Exception as e:
            st.error(f"Sistem Hatası: {str(e)}")

# --- 💬 SABİT İLETİŞİM & YAPIMCI BUTONU (SAYFA ALTI) ---
st.write("")
st.divider()
st.link_button(
    label=DIL_PAKETI[aktif_dil]['contact_btn'],
    url="https://threads.com@muratsenr",
    use_container_width=True
)
