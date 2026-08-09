import json
import re
import io
import zipfile
import urllib.parse
from pathlib import Path
from typing import Any, List, Dict
from datetime import datetime
import streamlit as st
import xlsxwriter

# --- MOBİL VE GENİŞ EKRAN GÜVENLİ STANDART AYARI ---
st.set_page_config(
    page_title="Threads Profil Takip Sistemi",
    page_icon="🎯",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# --- 📱 REAL PWA (ANA EKRANA EKLE) MOBİL UYGULAMA ENJEKTÖRÜ ---
st.components.v1.html(
    """
    <script>
    if ('serviceWorker' in navigator) {
      window.addEventListener('load', function() {
        navigator.serviceWorker.register('https://jsdelivr.net');
      });
    }
    </script>
    <link rel="manifest" href='data:application/manifest+json,{"short_name":"ThreadsTakip","name":"Threads Profil Takip Sistemi","icons":[{"src":"https://flaticon.com","type":"image/png","sizes":"512x512"}],"start_url":".","background_color":"#000000","theme_color":"#000000","display":"standalone","orientation":"portrait"}'>
    """,
    height=0,
)
# --- ÇOKLU DİL SÖZLÜĞÜ (TR / EN / DE) ---
DIL_PAKETI = {
    "TR": {
        "main_title": "THREADS TÜRKİYE / MURAT & ESRA İŞ BİRLİĞİ (CAN&KAN)",
        "main_sub": "YEREL VE GÜVENLİ ÇİFT YÖNLÜ PROFİL TAKİPÇİ SİSTEMİ // PWA MOBİL SÜRÜM ACTIVE",
        "main_hashtag": "#bendeğilbizyaptık",
        "load_batch": "📁 Ham .zip Dosyasını veya .json Dosyalarını Sürükleyin / Seçin",
        "btn_analyze": "ANALİZİ BAŞLAT",
        "tab_unfollowers": "Beni Takip Etmeyenler",
        "tab_fans": "Geri Takip Etmediklerim",
        "tab_ghosts": "Hayalet (Ghost) Hesaplar",
        "input_error_msg": "Analiz için veri havuzunda hem 'followers.json' hem de 'following.json' bulunmalıdır. Lütfen doğru dosya veya .zip yükleyin.",
        "parse_error_msg": "Yüklenen JSON şeması siber motor tarafından çözümlenemedi.",
        "success_msg": "Dosyalar tarandı, Excel ve TXT raporları köprü linkleriyle üretildi!",
        "perfect_sync": "🎉 [KUSURSUZ SENKRONİZASYON]: Herkes sizi geri takip ediyor!",
        "no_fans": "🎯 [HAYRAN YOK]: Takip ettiğiniz herkesi siz de geri takip ediyorsunuz.",
        "no_ghosts": "🛡️ [TEMİZ PROFİL]: Profilinizde hayalet veya bot hesap algılanmadı.",
        "download_excel": "📥 Excel Analiz Raporunu İndir",
        "summary_title": "📊 PROFİL SAĞLIK ÖZETİ",
        "health_score": "Sağlık Skoru",
        "guide_title": "📖 Threads Verileri Nasıl İndirilir? (Kullanım Kılavuzu)",
        "guide_step1": "📱 **%100 GÜVENLİ HİBRİT MOTOR:** İster ham .zip atın, ister içindeki .json dosyalarını çoklu seçip yükleyin.",
        "guide_step2": "1️⃣ **Instagram/Threads** uygulamasını açın ve **Ayarlar -> Hesaplar Merkezi** bölümüne girin.",
        "guide_step3": "2️⃣ **Bilgilerin ve İzinlerin -> Bilgilerini İndir** adımlarını takip edin.",
        "guide_step4": "3️⃣ **Indirme Talep Et** butonuna basın ve sadece **Threads** seçeneğini işaretleyin.",
        "guide_step5": "4️⃣ Dosya formatını **JSON** (ÖNEMLİ!), medya kalitesini **Düşük** seçip talebi onaylayın.",
        "guide_step6": "5️⃣ E-postanıza gelen ham `.zip` dosyasını klasöre açmadan direkt yükleyebilir veya içindeki `connections/followers_and_following` klasöründen dosyaları seçebilirsiniz.",
        "contact_btn": "💬 YAPIMCI İLE İLETİŞİME GEÇ (@muratsenr)",
        "player_title": "🎵 Arka Plan Müziğini Zorla Koydurdu",
        "search_placeholder": "🔍 Listede kullanıcı adı ara...",
        "chart_title": "📈 Profil Dağılım Grafiği",
        "sort_label": "⏳ Zaman Sıralaması",
        "sort_newest": "Önce En Yeni (Kronolojik)",
        "sort_oldest": "Önce En Eski (Nostaljik)",
        "history_title": "⏳ ŞU ANDA AKTİF DEĞİŞİM KARŞILAŞTIRICISI",
        "login_header": "🔒 SİSTEME GÜVENLİ GİRİŞ",
        "login_user": "Kullanıcı Adı",
        "login_pass": "Şifre",
        "login_btn": "GİRİŞ YAP",
        "login_success": "🔓 Erişim Yetkisi Onaylandı! Sistem yükleniyor...",
        "login_error": "❌ Hatalı Kullanıcı Adı veya Şifre! Lütfen tekrar deneyin.",
        "outdated_warning": "⏳ **DİKKAT: ESKİ VERİ SETİ ALGILANDI**\n\nYüklediğiniz veri paketleri en son {days} gün önce güncellenmiş görünüyor. En doğru sonuçlar için lütfen Threads verilerinizi yeniden indirin.",
        "logout_btn": "🔒 OTURUMU GÜVENLİ KAPAT (LOG-OUT)"
    },
    "EN": {
        "main_title": "THREADS GLOBAL/ MURAT & ESRA İŞ BİRLİĞİ",
        "main_sub": "LOCAL AND SECURE BI-DIRECTIONAL PROFILE ANALYSIS SYSTEM // PWA MOBILE ACTIVE",
        "main_hashtag": "#notmebutwe",
        "load_batch": "📁 Drag & Drop Raw .zip or Multiple .json Files Here",
        "btn_analyze": "START ANALYSIS",
        "tab_unfollowers": "Not Following Me Back",
        "tab_fans": "I Am Not Following Back",
        "tab_ghosts": "Ghost / Inactive Accounts",
        "input_error_msg": "Analysis requires both 'followers.json' and 'following.json' in the pool. Please upload correct files or .zip.",
        "parse_error_msg": "The uploaded JSON schema could not be resolved by the engine.",
        "success_msg": "Files scanned, Excel and TXT reports generated with clickable links!",
        "perfect_sync": "🎉 [PERFECT SYNC]: Everyone is following you back!",
        "no_fans": "🎯 [NO FANS]: You are following back everyone who follows you.",
        "no_ghosts": "🛡️ [CLEAN PROFILE]: No ghost or bot accounts detected on your profile.",
        "download_excel": "📥 Download Excel Analysis Report",
        "summary_title": "📊 PROFILE HEALTH SUMMARY",
        "health_score": "Health Score",
        "guide_title": "📖 How to Download Threads Data? (User Guide)",
        "guide_step1": "📱 **%100 SECURE HYBRID ENGINE:** Drag raw .zip or multiple select .json files as you wish.",
        "guide_step2": "1️⃣ Open **Instagram/Threads**, go to **Settings -> Accounts Center**.",
        "guide_step3": "2️⃣ Follow **Your Information and Permissions -> Download Your Information**.",
        "guide_step4": "3️⃣ Click **Request a Download** and select only **Threads**.",
        "guide_step5": "4️⃣ Choose format as **JSON**, media quality as **Low** and submit.",
        "guide_step6": "5️⃣ Upload the raw `.zip` file from your email directly, or browse and extract files inside connections folder.",
        "contact_btn": "💬 CONTACT DEVELOPER (@muratsenr)",
        "player_title": "🎵 Background Music: Cankan - Yaranamadım",
        "search_placeholder": "🔍 Search username in list...",
        "chart_title": "📈 Profile Distribution Chart",
        "sort_label": "⏳ Time Sorting",
        "sort_newest": "Newest First (Chronological)",
        "sort_oldest": "Oldest First",
        "history_title": "⏳ ACTIVE LIVE CHANGE COMPARATOR",
        "login_header": "🔒 SECURE SYSTEM LOGIN",
        "login_user": "Username",
        "login_pass": "Password",
        "login_btn": "LOGIN",
        "login_success": "🔓 Access Granted! Loading system...",
        "login_error": "❌ Invalid Username or Password! Please try again.",
        "outdated_warning": "⏳ **WARNING: OUTDATED DATA DETECTED**\n\nYour uploaded data files were last updated {days} days ago. For the most accurate results, please re-download from Threads.",
        "logout_btn": "🔒 SECURE LOG-OUT"
    },
    "DE": {
        "main_title": "THREADS GLOBAL",
        "main_sub": "LOKALES UND SICHERES BIDIREKTIONALES PROFIL-ANALYSESYSTEM // PWA MOBILE AKTIV",
        "main_hashtag": "#nichtichsondernwir",
        "load_batch": "📁 Ziehen Sie die unentpackte .zip oder mehrere .json Dateien hierher",
        "btn_analyze": "ANALYSE STARTEN",
        "tab_unfollowers": "Folgen mir nicht zurück",
        "tab_fans": "Ich folge nicht zurück",
        "tab_ghosts": "Geister / Inaktive Konten",
        "input_error_msg": "Die Analyse erfordert sowohl die Datei 'followers.json' als auch 'following.json' im Verzeichnis.",
        "parse_error_msg": "Das hochgeladene JSON-Schema konnte nicht aufgelöst werden.",
        "success_msg": "Dateien erfolgreich gescannt, Berichte exportiert!",
        "perfect_sync": "🎉 [SAFE LOG]: Jeder folgt Ihnen zurück!",
        "no_fans": "🎯 [SAFE LOG]: Sie folgen jedem zurück, der Ihnen folgt.",
        "no_ghosts": "🛡️ [SAFE LOG]: Keine Geister- oder Bot-Konten auf Ihrem Profil.",
        "download_excel": "📥 Excel-Analysebericht herunterladen",
        "summary_title": "📊 PROFIL-GESUNDHEITSÜBERSICHT",
        "health_score": "Gesundheitsscore",
        "guide_title": "📖 Wie lade ich Threads-Daten herunter? (Handbuch)",
        "guide_step1": "📱 **%100 SICHERE HYBRID-ENGINE:** Laden Sie entweder die rohe .zip-Datei oder einzelne JSON-Dateien hoch.",
        "guide_step2": "1️⃣ Öffnen Sie **Instagram/Threads**, gehen Sie zu **Einstellungen -> Kontenübersicht**.",
        "guide_step3": "2️⃣ Folgen Sie **Deine Informationen und Berechtigungen -> Deine Informationen herunterladen**.",
        "guide_step4": "3️⃣ Klicken Sie auf **Download anfordern** und wählen Sie nur **Threads** aus.",
        "guide_step5": "4️⃣ Wählen Sie das Format **JSON** und die Medienqualität **Niedrig**.",
        "guide_step6": "5️⃣ Entpacken Sie die erhaltene `.zip`-Datei auf Ihrem Gerät.",
        "guide_step7": "6️⃣ Ziehen Sie den Ordner `connections/followers_and_following` hierher.",
        "contact_btn": "💬 CYBER-ENTWICKLER KONTAKTIEREN (@muratsenr)",
        "player_title": "🎵 Hintergrundmusik: Cankan - Yaranamadım",
        "search_placeholder": "🔍 Suchen Sie nach Benutzernamen...",
        "chart_title": "📈 Profil-Verteilungsdiagramm",
        "sort_label": "⏳ Zeitliche Sortierung",
        "sort_newest": "Neueste zuerst",
        "sort_oldest": "Älteste zuerst",
        "history_title": "⏳ AKTIVER LIVE-VERÄNDERUNGSVERGLEICHER",
        "login_header": "🔒 SICHERER SYSTEM-LOGIN",
        "login_user": "Benutzername",
        "login_pass": "Passwort",
        "login_btn": "EINLOGGEN",
        "login_success": "🔓 Zugriff gewährt! System wird geladen...",
        "login_error": "❌ Ungültiger Benutzername oder Passwort! Bitte versuchen Sie es erneut.",
        "outdated_warning": "⏳ **WARNUNG: VERALTETE DATEN ERKANNT**\n\nIhre Datendateien wurden vor {days} Tagen aktualisiert. Bitte laden Sie Ihre Daten neu herunter.",
        "logout_btn": "🔒 SICHERER LOG-OUT"
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
# --- 🔑 ÖZEL KULLANICI ADI VE ŞİFRE VERİ TABANI ---
if "user_db" not in st.session_state:
    st.session_state.user_db = {
        "murat": "esra",
        "esra": "murat",
        "demo": "threads2026",
        "ömür": "deniz",
        "deniz": "ömür",
        "rıdvan": "gönül"
    }

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.markdown(f"### 🎯 Threads Profil Takip Sistemi")
    st.write("")
    
    with st.container(border=True):
        st.markdown("<h4 style='color:#3a7ebf; margin-top:0px;'><b>🔒 SİSTEME GÜVENLİ GİRİŞ</b></h4>", unsafe_allow_html=True)
        input_user = st.text_input("👤 Kullanıcı Adı", key="login_username_field").strip()
        input_pass = st.text_input("🔑 Şifre", type="password", key="login_password_field").strip()
        login_click = st.button("SİSTEME GİRİŞ YAP", use_container_width=True, type="primary")
        
        if login_click:
            if input_user in st.session_state.user_db and st.session_state.user_db[input_user] == input_pass:
                st.session_state.logged_in = True
                st.session_state.current_active_user = input_user
                st.success("🔓 Erişim Onaylandı! Sistem yükleniyor...")
                st.rerun()
            else:
                st.error("❌ Hatalı Kullanıcı Adı veya Şifre! Lütfen bilgilerinizi kontrol edin.")
    st.stop()

# --- TEMA SEÇİCİ (Hata Veren Satır Düzeltildi) ---
col_theme, col_space = st.columns(2)
with col_theme:
    tema_secimi = st.selectbox("🌓 Tema Modu", ["Karanlık Gece Modu", "Açık Threads Modu"], label_visibility="collapsed")

if tema_secimi == "Açık Threads Modu":
    st.markdown("""
        <style>
        .stApp { background-color: #ffffff !important; color: #000000 !important; }
        h1, h2, h3, h4, h5, h6, p, label, span { color: #000000 !important; }
        div[data-testid="stExpander"], div[data-testid="stFileUploader"] { background-color: #f3f4f6 !important; border: 1px solid #e5e7eb !important; }
        </style>
        """, unsafe_allow_html=True)

# --- ANA ARAYÜZ ---
st.title("🎯 Threads Profil Takip Sistemi")

col_lang, col_hashtag = st.columns(2)
with col_lang:
    aktif_dil = st.selectbox("🌐 Language / Dil", ["TR", "EN", "DE"])

with col_hashtag:
    st.markdown(f"<h4 style='text-align: right; color: #3a7ebf; margin-top: 5px;'>{DIL_PAKETI[aktif_dil]['main_hashtag']}</h4>", unsafe_allow_html=True)

st.markdown(f"### {DIL_PAKETI[aktif_dil]['main_title']}")
st.caption(DIL_PAKETI[aktif_dil]['main_sub'])
st.divider()

st.caption(DIL_PAKETI[aktif_dil]['player_title'])
st.link_button(label="▶️ YAPARKEN DİNLERSİNİZ YA (Göndermeli Şarkı)", url="https://youtube.com", use_container_width=True)

with st.expander(DIL_PAKETI[aktif_dil]['guide_title'], expanded=False):
    st.info(DIL_PAKETI[aktif_dil]['guide_step1'])
    st.markdown(DIL_PAKETI[aktif_dil]['guide_step2'])
    st.markdown(DIL_PAKETI[aktif_dil]['guide_step3'])
    st.markdown(DIL_PAKETI[aktif_dil]['guide_step4'])
    st.markdown(DIL_PAKETI[aktif_dil]['guide_step5'])
    st.markdown(DIL_PAKETI[aktif_dil]['guide_step6'])

st.write("") 

# --- 📁 HİBRİT DOSYA PANELİ (ZIP + JSON SÜRÜKLE / SEÇ) ---
uploaded_inputs = st.file_uploader(DIL_PAKETI[aktif_dil]['load_batch'], type=["zip", "json"], accept_multiple_files=True)

following_bytes = None
followers_bytes = None

if uploaded_inputs:
    for item in uploaded_inputs:
        if item.name.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(item) as z:
                    for file_info in z.infolist():
                        if "following.json" in file_info.filename.lower():
                            following_bytes = z.read(file_info.filename)
                        elif "followers.json" in file_info.filename.lower():
                            followers_bytes = z.read(file_info.filename)
            except Exception:
                st.error("Zip Çözümleme Hatası! Lütfen geçerli bir Threads arşivi yükleyin.")
        elif item.name.lower().endswith(".json"):
            if "following" in item.name.lower():
                following_bytes = item.read()
            elif "followers" in item.name.lower():
                followers_bytes = item.read()

btn_trigger = st.button(DIL_PAKETI[aktif_dil]['btn_analyze'], use_container_width=True, type="primary")
if btn_trigger or st.session_state.get('analyzed', False):
    if not following_bytes or not followers_bytes:
        st.warning(DIL_PAKETI[aktif_dil]['input_error_msg'])
    else:
        try:
            if 'unfollowers' not in st.session_state:
                if 'current_unfollowers' in st.session_state:
                    st.session_state.prev_unfollowers = st.session_state.current_unfollowers
                    st.session_state.prev_fans = st.session_state.current_fans
                    st.session_state.has_history = True
                else:
                    st.session_state.has_history = False

                following_raw = json.loads(following_bytes.decode("utf-8"))
                st.session_state.global_following_map = AnalizMotoru.akilli_süre_ayristir(following_raw)
                
                followers_raw = json.loads(followers_bytes.decode("utf-8"))
                st.session_state.global_followers_map = AnalizMotoru.akilli_süre_ayristir(followers_raw)
                    
                st.session_state.following_set = set(st.session_state.global_following_map.keys())
                st.session_state.followers_set = set(st.session_state.global_followers_map.keys())
                
                st.session_state.current_unfollowers = st.session_state.following_set - st.session_state.followers_set
                st.session_state.current_fans = st.session_state.followers_set - st.session_state.following_set
                st.session_state.ghosts = {u for u, ts in st.session_state.global_followers_map.items() if AnalizMotoru.bot_ve_pasiflik_kontrolü(u, ts)}
                st.session_state.analyzed = True

            global_following_map = st.session_state.global_following_map
            global_followers_map = st.session_state.global_followers_map
            following_set = st.session_state.following_set
            followers_set = st.session_state.followers_set
            unfollowers = st.session_state.current_unfollowers
            fans = st.session_state.current_fans
            ghosts = st.session_state.ghosts

            if not following_set or not followers_set:
                st.error(DIL_PAKETI[aktif_dil]['parse_error_msg'])
            else:
                # --- AKILLI DOSYA GÜNCELLİK DENETLEYİCİSİ SÜZGECİ ---
                en_son_sinyal_zamani = max(list(global_following_map.values()) + list(global_followers_map.values()), default=0)
                if en_son_sinyal_zamani > 0:
                    gecen_gun = (datetime.now() - datetime.fromtimestamp(en_son_sinyal_zamani)).days
                    if gecen_gun > 30:
                        st.warning(DIL_PAKETI[aktif_dil]['outdated_warning'].replace("{days}", str(gecen_gun)))

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

                # --- DEĞİŞİM KARŞILAŞTIRICISI PANELİ ---
                if st.session_state.get('has_history', False):
                    st.write("")
                    st.markdown(f"##### {DIL_PAKETI[aktif_dil]['history_title']}")
                    yeni_takipten_cikanlar = unfollowers - st.session_state.prev_unfollowers
                    yeni_hayranlar = fans - st.session_state.prev_fans
                    
                    c1, c2 = st.columns(2)
                    c1.metric("🚨 Yeni Taktikten Çıkanlar", len(yeni_takipten_cikanlar), f"+{len(yeni_takipten_cikanlar)} Kişi" if yeni_takipten_cikanlar else "Değişim Yok", delta_color="inverse")
                    c2.metric("🛸 Yeni Kazanılan Hayranlar", len(yeni_hayranlar), f"+{len(yeni_hayranlar)} Kişi" if yeni_hayranlar else "Değişim Yok")

                # --- 📊 YEREL SÜTUN GRAFİĞİ ENTEGRASYONU (Hizalaması Düzeltildi) ---
                st.write("")
                st.markdown(f"##### {DIL_PAKETI[aktif_dil]['chart_title']}")
                chart_data = {
                    "Kategori": ["Takip Etmeyenler", "Karşılıklı", "Hayranlar", "Hayaletler"],
                    "Sayı": [len(unfollowers), len(following_set & followers_set), len(fans), len(ghosts)]
                }
                st.bar_chart(data=chart_data, x="Kategori", y="Sayı", use_container_width=True)
                st.write("")

                # --- BELLEKTE EXCEL OLUŞTURMA MOTORU ---
                output_excel = io.BytesIO()
                workbook = xlsxwriter.Workbook(output_excel)
                header_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#1f4e78', 'border': 1, 'align': 'center'})
                link_format = workbook.add_format({'font_color': 'blue', 'underline': True})
                text_format = workbook.add_format({'align': 'left'})
                sheet_unf = workbook.add_worksheet("Beni Takip Etmeyenler")
                sheet_unf.write_row('A1', ['No', 'Kullanıcı Adı', 'Profil Linki', 'Süre'], header_format)
                sorted_unf_excel = sorted(unfollowers, key=lambda x: global_following_map.get(x, 0))
                for idx, user in enumerate(sorted_unf_excel, 1):
                    süre = AnalizMotoru.zaman_metnine_cevir(global_following_map.get(user, 0))
                    p_url = f"https://threads.com@{user}"
                    sheet_unf.write(idx, 0, idx)
                    sheet_unf.write(idx, 1, f"@{user}", text_format)
                    sheet_unf.write_url(idx, 2, p_url, link_format, string=p_url)
                    sheet_unf.write(idx, 3, süre, text_format)
                sheet_unf.set_column('A:A', 5); sheet_unf.set_column('B:B', 20); sheet_unf.set_column('C:C', 45); sheet_unf.set_column('D:D', 20)

                sheet_fans = workbook.add_worksheet("Geri Takip Etmediklerim")
                sheet_fans.write_row('A1', ['No', 'Kullanıcı Adı', 'Profil Linki', 'Süre'], header_format)
                sorted_fans_excel = sorted(fans, key=lambda x: global_followers_map.get(x, 0))
                for idx, user in enumerate(sorted_fans_excel, 1):
                    süre = AnalizMotoru.zaman_metnine_cevir(global_followers_map.get(user, 0))
                    p_url = f"https://threads.com@{user}"
                    sheet_fans.write(idx, 0, idx)
                    sheet_fans.write(idx, 1, f"@{user}", text_format)
                    sheet_fans.write_url(idx, 2, p_url, link_format, string=p_url)
                    sheet_fans.write(idx, 3, süre, text_format)
                sheet_fans.set_column('A:A', 5); sheet_fans.set_column('B:B', 20); sheet_fans.set_column('C:C', 45); sheet_fans.set_column('D:D', 20)

                sheet_gh = workbook.add_worksheet("Hayalet Hesaplar")
                sheet_gh.write_row('A1', ['No', 'Kullanıcı Adı', 'Profil Linki', 'Süre'], header_format)
                sorted_gh_excel = sorted(ghosts, key=lambda x: global_followers_map.get(x, 0))
                for idx, user in enumerate(sorted_gh_excel, 1):
                    süre = AnalizMotoru.zaman_metnine_cevir(global_followers_map.get(user, 0))
                    p_url = f"https://threads.com@{user}"
                    sheet_gh.write(idx, 0, idx)
                    sheet_gh.write(idx, 1, f"@{user}", text_format)
                    sheet_gh.write_url(idx, 2, p_url, link_format, string=p_url)
                    sheet_gh.write(idx, 3, süre, text_format)
                sheet_gh.set_column('A:A', 5); sheet_gh.set_column('B:B', 20); sheet_gh.set_column('C:C', 45); sheet_gh.set_column('D:D', 20)
                workbook.close()
                output_excel.seek(0)
                
                st.download_button(label=DIL_PAKETI[aktif_dil]['download_excel'], data=output_excel, file_name="Threads_Detayli_Analiz_Raporu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

                # --- KRONOLOJİK ZAMAN MOTORU ---
                st.markdown(f"###### {DIL_PAKETI[aktif_dil]['sort_label']}")
                sort_choice = st.radio(label="", options=[DIL_PAKETI[aktif_dil]['sort_newest'], DIL_PAKETI[aktif_dil]['sort_oldest']], horizontal=True, label_visibility="collapsed")
                is_reverse = (sort_choice == DIL_PAKETI[aktif_dil]['sort_newest'])

                # --- CANLI ARAMA ---
                search_query = st.text_input("", placeholder=DIL_PAKETI[aktif_dil]['search_placeholder']).strip().lower()
                clean_query = search_query.replace("@", "")

                sorted_unf = sorted(unfollowers, key=lambda x: global_following_map.get(x, 0), reverse=is_reverse)
                sorted_fans = sorted(fans, key=lambda x: global_followers_map.get(x, 0), reverse=is_reverse)
                sorted_gh = sorted(ghosts, key=lambda x: global_followers_map.get(x, 0), reverse=is_reverse)

                t1, t2, t3 = st.tabs([DIL_PAKETI[aktif_dil]['tab_unfollowers'], DIL_PAKETI[aktif_dil]['tab_fans'], DIL_PAKETI[aktif_dil]['tab_ghosts']])
                with t1:
                    filtered_unf = [u for u in sorted_unf if clean_query in u.lower()] if clean_query else sorted_unf
                    if filtered_unf:
                        for index, user in enumerate(filtered_unf, 1):
                            süre = AnalizMotoru.zaman_metnine_cevir(global_following_map.get(user, 0))
                            st.markdown(f"[{index:03d}] 🔗 [@{user}](https://threads.com/@{user}) &nbsp;&nbsp;&nbsp;&nbsp; <b>⌛ {süre}</b>", unsafe_allow_html=True)
                    else: st.info(DIL_PAKETI[aktif_dil]['perfect_sync'] if not clean_query else "Eşleşen kullanıcı bulunamadı.")
                with t2:
                    filtered_fans = [u for u in sorted_fans if clean_query in u.lower()] if clean_query else sorted_fans
                    if filtered_fans:
                        for index, user in enumerate(filtered_fans, 1):
                            süre = AnalizMotoru.zaman_metnine_cevir(global_followers_map.get(user, 0))
                            st.markdown(f"[{index:03d}] 🔗 [@{user}](https://threads.com/@{user}) &nbsp;&nbsp;&nbsp;&nbsp; <b>⌛ {süre}</b>", unsafe_allow_html=True)
                    else: st.info(DIL_PAKETI[aktif_dil]['no_fans'] if not clean_query else "Eşleşen kullanıcı bulunamadı.")
                with t3:
                    filtered_gh = [u for u in sorted_gh if clean_query in u.lower()] if clean_query else sorted_gh
                    if filtered_gh:
                        for index, user in enumerate(filtered_gh, 1):
                            süre = AnalizMotoru.zaman_metnine_cevir(global_followers_map.get(user, 0))
                            st.markdown(f"[{index:03d}] 🔗 [@{user}](https://threads.com/@{user}) &nbsp;&nbsp;&nbsp;&nbsp; <b>⌛ {süre}</b>", unsafe_allow_html=True)
                    else: st.info(DIL_PAKETI[aktif_dil]['no_ghosts'] if not clean_query else "Eşleşen kullanıcı bulunamadı.")
        except Exception as e: st.error(f"Sistem Hatası: {str(e)}")

# --- LOG-OUT VE SİSTEM AYARLARI ALT ALANI ---
st.write(""); st.divider()

with st.expander("⚙️ Hesap Ayarları (Şifre Değiştir)", expanded=False):
    yeni_sifre_input = st.text_input("🔑 Yeni Şifrenizi Girin", type="password", key="change_password_box").strip()
    sifre_onay_btn = st.button("ŞİFREYİ GÜNCELLE", use_container_width=True)
    if sifre_onay_btn and yeni_sifre_input:
        st.session_state.user_db[st.session_state.current_active_user] = yeni_sifre_input
        st.success("🎉 Şifreniz başarıyla güncellendi! Bir sonraki girişte aktif olacaktır.")

st.write("")
logout_click = st.button(DIL_PAKETI[aktif_dil]['logout_btn'], use_container_width=True, type="secondary")
if logout_click:
    st.session_state.logged_in = False
    st.session_state.analyzed = False
    st.rerun()

st.write("")
st.link_button(label=DIL_PAKETI[aktif_dil]['contact_btn'], url="https://threads.com/@muratsenr", use_container_width=True)
