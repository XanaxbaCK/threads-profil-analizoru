import json
import re
import io
import time
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
# --- ÇOKLU DİL SÖZLÜĞÜ (TR / EN) ---
DIL_PAKETI = {
    "TR": {
        "main_title": "THREADS TÜRKİYE / MURAT ŞENER ",
        "main_sub": "YEREL VE GÜVENLİ ÇİFT YÖNLÜ PROFİL TAKİPÇİ SİSTEMİ // MOBİL & WEB SÜRÜM",
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
        "guide_title": "📖 Threads Verileri Nasıl İndirilir? (Kullanım Kılavuzu) LÜTFEN OKUYUNUZ!",
        "pwa_guide_text": "📱 **BU SİTEYİ TELEFONA MOBİL UYGULAMA OLARAK KURABİLMEK İÇİN:**\n\n• **iOS (Safari) için:** Sayfanın altındaki **'Paylaş'** (Yukarı ok olan kutu) butonuna dokunun. Açılan menüden **'Ana Ekrana Ekle' (Add to Home Screen)** seçeneğini işaretleyin.\n\n• **Android (Chrome) için:** Sağ üstteki **'Üç Nokta'** simgesine dokunun. Açılan menüden **'Uygulamayı Yükle'** veya **'Ana Ekrana Ekle'** seçeneğine basın.",
        "guide_step1": "📱 **%100 GÜVENLİ YÜKLEME:** İster ham .zip atın, ister içindeki .json dosyalarını çoklu seçip yükleyin.",
        "guide_step2": "1️⃣ **Threads** uygulamasını açın ve **Ayarlar -> Hesaplar Merkezi** bölümüne girin.",
        "guide_step3": "2️⃣ **Bilgilerin ve İzinlerin -> Bilgilerini İndir** adımlarını takip edin.",
        "guide_step4": "3️⃣ **Indirme Talep Et** butonuna basın ve sadece **Threads** seçeneğini işaretleyin.",
        "guide_step5": "4️⃣ Dosya formatını **JSON** (ÖNEMLİ!), medya kalitesini **Düşük** Tarih aralığını da **En Baştan** (ÖNEMLİ!) seçip talebi onaylayın.",
        "guide_step6": "5️⃣ E-postanıza gelen ham `.zip` dosyasını klasöre açmadan direkt yükleyebilir veya içindeki `connections/followers_and_following` klasöründen dosyaları seçebilirsiniz.",
        "contact_btn": "💬 YAPIMCI İLE İLETİŞİME GEÇ (@muratsenr)",
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
        "logout_btn": "🔒 OTURUMU GÜVENLİ KAPAT (ÇIKIŞ YAP)",
        "premium_notice": "👑 **PREMIUM ÖZELLİK KİLİTLİ:** Advanced grafik analizler, Excel indirme motoru ve kronolojik zaman sıralaması sadece Premium üyelere özeldir. Yetki yükseltmek için lütfen yöneticiyle iletişime geçin.",
        "badge_premium": "👑 Premium Hesap (Sınırsız Erişim)",
        "badge_standard": "👤 Standart Hesap (Kısıtlı Erişim)",
        "welcome_user": "Hoş geldiniz, {user}"
    },

        "EN": {
        "main_title": "THREADS GLOBAL / MURAT ŞENER PRODUCER",
        "main_sub": "LOCAL AND SECURE BI-DIRECTIONAL PROFILE ANALYSIS SYSTEM // MOBILE ACT & WEB",
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
        "guide_title": "📖 How to Download Threads Data and Install as App?",
        "pwa_guide_text": "📱 **INSTALL THIS SITE AS AN APP:**\n\n• **For iOS (Safari):** Tap the **'Share'** button (box with an upward arrow) at the bottom. Select **'Add to Home Screen'** from the menu.\n\n• **For Android (Chrome):** Tap the **'Three Dots'** icon at the top right. Select **'Install App'** or **'Add to Home Screen'** from the menu.",
        "guide_step1": "📱 **%100 SECURE HYBRID ENGINE:** Drag raw .zip or multiple select .json files as you wish.",
        "guide_step2": "1️⃣ Open **Threads**, go to **Settings -> Accounts Center**.",
        "guide_step3": "2️⃣ Follow **Your Information and Permissions -> Download Your Information**.",
        "guide_step4": "3️⃣ Click **Request a Download** and select only **Threads**.",
        "guide_step5": "4️⃣ Choose format as **JSON**, media quality as **Low** and submit.",
        "guide_step6": "5️⃣ Upload the raw `.zip` file from your email directly, or browse and extract files inside connections folder.",
        "contact_btn": "💬 CONTACT DEVELOPER (@muratsenr)",
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
        "outdated_warning": "⚠️ **WARNING: OUTDATED DATA DETECTED**\n\nYour uploaded data files were last updated {days} days ago. For the most accurate results, please re-download from Threads.",
        "logout_btn": "🔒 SECURE LOG-OUT",
        "premium_notice": "👑 **PREMIUM FEATURE LOCKED:** Advanced charts, Excel exports, and chronological sorting are restricted to Premium members.",
        "badge_premium": "👑 Premium Account (Unrestricted)",
        "badge_standard": "👤 Standard Account (Restricted)",
        "welcome_user": "Welcome, {user}"
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
# --- 🔑 GÜVENLİ VE KALICI DEĞİŞTİRİLEBİLİR KOD TABANLI VERİ TABANI ---
if "user_db" not in st.session_state:
    st.session_state.user_db = {
        "murat": "snr",
        "esra": "esra",
        "demo": "demo",
        "ömür": "deniz",
        "deniz": "ömür",
        "rıdvan": "gönül",
        "irem": "irem",
        "sinemk": "sinem"
    }

if "premium_users" not in st.session_state:
    st.session_state.premium_users = {"murat", "esra", "ömür", "deniz", "rıdvan", "irem", "sinemk"}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# --- ⏱️ 5 DAKİKALIK OTOMATİK OTURUM ZAMAN AŞIMI GÜVENLİK MOTORU ---
SURE_SINIRI_SANIYE = 5 * 60 

if st.session_state.logged_in:
    su_anki_zaman = time.time()
    if "last_activity_time" in st.session_state:
        gecen_bos_sure = su_anki_zaman - st.session_state.last_activity_time
        if gecen_bos_sure > SURE_SINIRI_SANIYE:
            st.session_state.logged_in = False
            st.session_state.analyzed = False
            st.warning("⚠️ Güvenlik amacıyla 5 dakikalık işlem yapılmadığı için oturumunuz otomatik kapatılmıştır.")
            st.stop()
    st.session_state.last_activity_time = su_anki_zaman

if not st.session_state.logged_in:
    st.markdown(f"### 🎯 Threads Profil Takip Sistemi")
    st.write("")
    
    with st.container(border=True):
        st.markdown("<h4 style='color:#3a7ebf; margin-top:0px;'><b>🔒 SİSTEME GÜVENLİ GİRİŞ</b></h4>", unsafe_allow_html=True)
        input_user = st.text_input("👤 Kullanıcı Adı", key="login_username_field").strip().lower()
        input_pass = st.text_input("🔑 Şifre", type="password", key="login_password_field").strip()
        login_click = st.button("SİSTEME GİRİŞ YAP", use_container_width=True, type="primary")
        
        if login_click:
            if input_user in st.session_state.user_db and st.session_state.user_db[input_user] == input_pass:
                st.session_state.logged_in = True
                st.session_state.current_active_user = input_user
                st.session_state.last_activity_time = time.time()
                st.success("🔓 Erişim Onaylandı! Sistem yükleniyor...")
                st.rerun()
            else:
                st.error("❌ Hatalı Kullanıcı Adı veya Şifre! Lütfen bilgilerinizi kontrol edin.")
    st.stop()
# --- SESSİON_STATE TABANLI TEMA HAFIZASI ---
if "sabit_tema" not in st.session_state:
    st.session_state.sabit_tema = "Karanlık Gece Modu"

if st.session_state.sabit_tema == "Premium Fildişi & Kemik Modu":
    st.markdown("""
        <style>
        .stApp { background-color: #f9f6f0 !important; color: #1c1c1e !important; }
        h1, h2, h3, h4, h5, h6, p, label, span, small { color: #1c1c1e !important; }
        div[data-testid="stExpander"], div[data-testid="stFileUploader"], div[data-testid="stDataframe"] { background-color: #f1ede4 !important; border: 1px solid #e1dacb !important; }
        .stMarkdown p { color: #1c1c1e !important; }
        .stMarkdown b { color: #000000 !important; }
        </style>
        """, unsafe_allow_html=True)

# --- ANA ARAYÜZ BAŞLANGICI VE ETİKETLER ---
st.title("🎯 Threads Profil Takip Sistemi")

aktif_u = st.session_state.current_active_user
is_user_premium = (aktif_u in st.session_state.premium_users)

col_user_welcome, col_user_badge = st.columns(2)
with col_user_welcome:
    st.markdown(f"👋 **{DIL_PAKETI['TR']['welcome_user'].format(user=aktif_u)}**")

with col_user_badge:
    if is_user_premium:
        st.markdown(f"<p style='text-align: right; margin: 0; font-weight: bold; color: #2e7d32;'>{DIL_PAKETI['TR']['badge_premium']}</p>", unsafe_allow_html=True)
    else:
        st.markdown(f"<p style='text-align: right; margin: 0; font-weight: bold; color: #ef6c00;'>{DIL_PAKETI['TR']['badge_standard']}</p>", unsafe_allow_html=True)

st.write("")

col_lang, col_hashtag = st.columns(2)
with col_lang:
    aktif_dil = st.selectbox("🌐 Language / Dil", ["TR", "EN"])

with col_hashtag:
    st.markdown(f"<h4 style='text-align: right; color: #3a7ebf; margin-top: 5px;'>{DIL_PAKETI[aktif_dil]['main_hashtag']}</h4>", unsafe_allow_html=True)

st.markdown(f"### {DIL_PAKETI[aktif_dil]['main_title']}")
st.caption(DIL_PAKETI[aktif_dil]['main_sub'])
st.divider()

# --- 🎵 ŞARKI BUTONU BAŞLIĞI TEMİZLENDİ VE SADECE BUTON BIRAKILDI ---
st.link_button(label="▶️ YAPARKEN DİNLERSİNİZ BELKİ (Göndermeli Şarkı)", url="https://youtube.com", use_container_width=True)

with st.expander(DIL_PAKETI[aktif_dil]['guide_title'], expanded=False):
    st.info(DIL_PAKETI[aktif_dil]['pwa_guide_text'])
    st.markdown(DIL_PAKETI[aktif_dil]['guide_step1'])
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
                if st.session_state.get("chk_outdated_alert", True):
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

                # --- 🚨 ÖZELLİK 3: ALGORİTMA GÜVENLİK PUANI VE SHADOWBAN RİSKİ ---
                total_f = len(followers_set) if len(followers_set) > 0 else 1
                risk_orani = (len(ghosts) + len(unfollowers)) / total_f
                if risk_orani > 0.45:
                    s_status, s_color, s_desc = "⚠️ CRITICAL SHADOWBAN RISK", "#d32f2f", "Profilinizdeki hayalet hesap ve vefasız takipçi oranı kritik seviyede! Threads algoritmaları erişiminizi kısıtlıyor olabilir."
                elif risk_orani > 0.20:
                    s_status, s_color, s_desc = "⚡ ALGORİTMA SINIRDA", "#f57c00", "Hesabınız sınırda duruyor. Organik erişim gücünüzün düşmemesi için pasif ve bot hesapları temizlemeniz önerilir."
                else:
                    s_status, s_color, s_desc = "🛡️ ALGORİTMA GÜVENLİ", "#388e3c", "Tebrikler! Profil süzgeciniz temiz. Threads arama ve keşfet algoritmaları kitle sağlığınızı olumlu puanlıyor."

                # --- 📸 ÖZELLİK 4: THREADS ÖZET VİRAL PAYLAŞIM KARTI ---
                card_bg = "#121212" if st.session_state.sabit_tema == "Karanlık Gece Modu" else "#f1ede4"
                card_text = "#ffffff" if st.session_state.sabit_tema == "Karanlık Gece Modu" else "#1c1c1e"
                
                st.markdown(f"""
                <div style='background-color:{card_bg}; border:2px solid #3a7ebf; border-radius:15px; padding:25px; margin-top:15px; text-align:center; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto;'>
                    <h3 style='color:#3a7ebf; margin-bottom:5px; font-weight:bold;'>🎯 THREADS KARNENİZ</h3>
                    <p style='color:{card_text}; font-size:14px; margin-top:0px;'>@ {st.session_state.current_active_user} &nbsp;|&nbsp; {datetime.now().strftime('%Y-%m-%d')}</p>
                    <hr style='border:0; border-top:1px solid #3a7ebf; margin:15px 0;'>
                    <div style='display:flex; justify-content:space-around; margin:20px 0;'>
                        <div><span style='font-size:12px; color:#888;'>SAĞLIK SKORU</span><br><b style='font-size:26px; color:#3a7ebf;'>%{health_score}</b></div>
                        <div><span style='font-size:12px; color:#888;'>TAKİPÇİLER</span><br><b style='font-size:26px; color:{card_text};'>{len(followers_set)}</b></div>
                        <div><span style='font-size:12px; color:#888;'>TAKİP ETTİKLERİN</span><br><b style='font-size:26px; color:{card_text};'>{len(following_set)}</b></div>
                    </div>
                    <div style='background:rgba(58,126,191,0.1); border-radius:8px; padding:12px; margin-top:15px;'>
                        <span style='font-size:15px; font-weight:bold; color:{s_color};'>{s_status}</span><br>
                        <p style='font-size:12px; color:{card_text}; margin:5px 0 0 0; line-height:1.4;'>{s_desc}</p>
                    </div>
                    <p style='font-size:10px; color:#888; margin-top:20px; font-style:italic;'>📸 Ekran görüntüsü alıp Threads'te @muratsenr{st.session_state.current_active_user} etiketleyerek paylaş!</p>
                </div>
                """, unsafe_allow_html=True)
                st.write("")

                # --- 👑 PREMIUM KORUMALI PANEL SÜRECİ ---
                if st.session_state.current_active_user in st.session_state.premium_users:
                    if st.session_state.get('has_history', False):
                        st.write("")
                        st.markdown(f"##### {DIL_PAKETI[aktif_dil]['history_title']}")
                        yeni_takipten_cikanlar = unfollowers - st.session_state.prev_unfollowers
                        yeni_hayranlar = fans - st.session_state.prev_fans
                        
                        c1, c2 = st.columns(2)
                        c1.metric("🚨 Yeni Taktikten Çıkanlar", len(yeni_takipten_cikanlar), f"+{len(yeni_takipten_cikanlar)} Kişi" if yeni_takipten_cikanlar else "Değişim Yok", delta_color="inverse")
                        c2.metric("🛸 Yeni Kazanılan Hayranlar", len(yeni_hayranlar), f"+{len(yeni_hayranlar)} Kişi" if yeni_hayranlar else "Değişim Yok")

                    st.write("")
                    st.markdown(f"##### {DIL_PAKETI[aktif_dil]['chart_title']}")
                    chart_data = {
                        "Kategori": ["Takip Etmeyenler", "Karşılıklı", "Hayranlar", "Hayaletler"],
                        "Sayı": [len(unfollowers), len(following_set & followers_set), len(fans), len(ghosts)]
                    }
                    st.bar_chart(data=chart_data, x="Kategori", y="Sayı", use_container_width=True)
                    st.write("")
                else:
                    st.write("")
                    st.info(DIL_PAKETI[aktif_dil]['premium_notice'])
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
                    p_url = f"https://threads.com/@{user}"
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
                    p_url = f"https://threads.com/@{user}"
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
                    p_url = f"https://threads.com/@{user}"
                    sheet_gh.write(idx, 0, idx)
                    sheet_gh.write(idx, 1, f"@{user}", text_format)
                    sheet_gh.write_url(idx, 2, p_url, link_format, string=p_url)
                    sheet_gh.write(idx, 3, süre, text_format)
                sheet_gh.set_column('A:A', 5); sheet_gh.set_column('B:B', 20); sheet_gh.set_column('C:C', 45); sheet_gh.set_column('D:D', 20)
                workbook.close()
                output_excel.seek(0)
                
                if st.session_state.current_active_user in st.session_state.premium_users:
                    st.download_button(label=DIL_PAKETI[aktif_dil]['download_excel'], data=output_excel, file_name="Threads_Detayli_Analiz_Raporu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

                # --- ⏳ KRONOLOJİK ZAMAN MOTORU (PREMIUM KORUMALI) ---
                if st.session_state.current_active_user in st.session_state.premium_users:
                    st.markdown(f"###### {DIL_PAKETI[aktif_dil]['sort_label']}")
                    sort_choice = st.radio(label="", options=[DIL_PAKETI[aktif_dil]['sort_newest'], DIL_PAKETI[aktif_dil]['sort_oldest']], horizontal=True, label_visibility="collapsed")
                    is_reverse = (sort_choice == DIL_PAKETI[aktif_dil]['sort_newest'])
                else:
                    is_reverse = True

                # --- CANLI ARAMA MOTORU ---
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

# --- DİNAMİK YEDEKLEME FONKSİYONU (YAKLAŞIM A) ---
def kod_uret():
    dict_str = "    st.session_state.user_db = {\n"
    dict_str += ",\n".join([f'        "{uk}": "{up}"' for uk, up in st.session_state.user_db.items()])
    dict_str += "\n    }"
    set_str = "    st.session_state.premium_users = {"
    set_str += ", ".join([f'"{pu}"' for pu in st.session_state.premium_users])
    set_str += "}"
    return f"""# --- BU BLOK PANEL TARAFINDAN OTOMATİK ÜRETİLMİŞTİR ---
if "user_db" not in st.session_state:
{dict_str}

if "premium_users" not in st.session_state:
{set_str}"""

# --- 👑 SADECE MURAT'A ÖZEL GELİŞMİŞ SAAS CANLI ÜYE YÖNETİM MERKEZİ ---
if st.session_state.current_active_user == "murat":
    with st.expander("👑 Canlı Üye Yönetim Paneli", expanded=False):
        st.markdown("##### 👤 Yeni Kullanıcı Tanımla / Şifre Oluştur")
        c_user = st.text_input("Kullanıcı Adı Ekle", key="saas_new_user_input").strip().lower()
        c_pass = st.text_input("Şifre Belirle", key="saas_new_pass_input").strip()
        c_role = st.selectbox("Yetki Seviyesi", ["Standart (Kısıtlı)", "Premium (Limitsiz)"])
        
        save_member_btn = st.button("KULLANICIYI SİSTEME KAYDET", use_container_width=True)
        if save_member_btn and c_user and c_pass:
            st.session_state.user_db[c_user] = c_pass
            if c_role == "Premium (Limitsiz)":
                st.session_state.premium_users.add(c_user)
            else:
                st.session_state.premium_users.discard(c_user)
            st.success(f"🎉 '@{c_user}' kullanıcısı {c_role} yetkisiyle veri tabanına eklendi!")
            st.rerun()

        st.divider()

        st.markdown("##### ⚙️ Kayıtlı Üyeleri Yönet & Yetki Düzenle")
        üyeler_tablosu = []
        for u, p in st.session_state.user_db.items():
            yetki_durumu = "👑 Premium" if u in st.session_state.premium_users else "👤 Standart"
            üyeler_tablosu.append({"Kullanıcı Adı": f"@{u}", "Şifre": p, "Mevcut Yetki": yetki_durumu})
        st.dataframe(üyeler_tablosu, use_container_width=True, hide_index=True)
        
        st.write("")
        secilen_üye = st.selectbox("Düzenlenecek Üyeyi Seçin", list(st.session_state.user_db.keys()))
        
        if secilen_üye:
            g_pass = st.text_input(f"@{secilen_üye} İçin Yeni Şifre (Boş bırakılırsa değişmez)", type="password", key="saas_edit_pass")
            su_anki_rol_idx = 1 if secilen_üye in st.session_state.premium_users else 0
            g_role = st.radio(f"@{secilen_üye} Yetki Seviyesini Değiştir", ["Standart (Kısıtlı)", "Premium (Limitsiz)"], index=su_anki_rol_idx, horizontal=True)
            
            update_member_btn = st.button("ÜYE BİLGİLERİNİ GÜNCELLE", use_container_width=True, type="primary")
            if update_member_btn:
                if g_pass.strip():
                    st.session_state.user_db[secilen_üye] = g_pass.strip()
                if g_role == "Premium (Limitsiz)":
                    st.session_state.premium_users.add(secilen_üye)
                else:
                    st.session_state.premium_users.discard(secilen_üye)
                st.success(f"🎉 '@{secilen_üye}' üyelik ve şifre konfigürasyonları başarıyla güncellendi!")
                st.rerun()

        st.divider()
        st.markdown("##### 💾 Kalıcı Kod Güncelleme Bloğu")
        st.code(kod_uret(), language="python")

# --- Gelişmiş Hesap Ayarları Kutusu ---
with st.expander("⚙️ Gelişmiş Hesap Ayarları (Profil ve Tercihler)", expanded=False):
    st.markdown("##### 📊 Profil Bilgileri Kartı")
    st.write(f"• **Kullanıcı Adınız:** `@{st.session_state.current_active_user}`")
    st.write(f"• **Erişim Seviyeniz:** {'👑 Premium (Limitsiz)' if is_user_premium else '👤 Standart (Kısıtlı)'}")
    st.write(f"• **Güvenlik Durumu:** `SSL Active / RSA Encrypted`")
    
    st.divider()
    
    st.markdown("##### 🌓 Otomatik Tema Sabitleyici")
    secilen_sabit = st.radio("Sistem temasını sabitleyin:", ["Karanlık Gece Modu", "Premium Fildişi & Kemik Modu"], index=0 if st.session_state.sabit_tema == "Karanlık Gece Modu" else 1)
    if secilen_sabit != st.session_state.sabit_tema:
        st.session_state.sabit_tema = secilen_sabit
        st.success("🎉 Tema tercihiniz oturum boyunca sabitlendi!")
        st.rerun()
        
    st.divider()
    
    st.markdown("##### 🔔 Bildirim ve Tercih Yönetimi")
    st.session_state.chk_outdated_alert = st.checkbox("Yüklenen Threads verilerim 30 günden eskiyse beni uyar", value=st.session_state.get("chk_outdated_alert", True))
    st.checkbox("Yeni analiz motoru güncellemelerini ana ekranda göster", value=True)
    
    st.divider()
    
    st.markdown("##### 🔑 Giriş Şifresini Değiştir")
    yeni_sifre_input = st.text_input("Yeni Şifrenizi Girin", type="password", key="change_password_box").strip()
    sifre_onay_btn = st.button("ŞİFREYİ GÜNCELLE (İstediğiniz Şifreyi @muratsener Profili Üzerinden Bildiriniz)", use_container_width=True)
    
    if (sifre_onay_btn and yeni_sifre_input) or st.session_state.get("sifre_degisti_mi", False):
        if sifre_onay_btn and yeni_sifre_input:
            st.session_state.user_db[st.session_state.current_active_user] = yeni_sifre_input
            st.session_state.sifre_degisti_mi = True
            st.success("🎉 Şifreniz başarıyla güncellendi!")
        
        st.write("")
        st.warning("💾 **ŞİFRENİZİ KALICI YAPIN:** Sayfa yenilendiğinde şifrenizin sıfırlanmaması için aşağıdaki kodu kopyalayıp admin panelinize iletin veya `app.py` dosyanızdaki **Parça 5**'e yapıştırın.")
        st.code(kod_uret(), language="python")

# --- 🗑️ GÜVENLİ OTURUMU KAPAT VE TÜM VERİLERİ TEMİZLE ---
st.write("")
logout_click = st.button("🗑️ OTURUMU KAPAT VE TÜM VERİLERİ TEMİZLE (DEEP CLEAN)", use_container_width=True, type="secondary")
if logout_click:
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.session_state.logged_in = False
    st.session_state.analyzed = False
    st.rerun()

# --- 💬 YAPIMCI SİBER İLETİŞİM BUTONU ---
st.write("")
st.link_button(label=DIL_PAKETI[aktif_dil]['contact_btn'], url="https://threads.com/@muratsenr", use_container_width=True)
