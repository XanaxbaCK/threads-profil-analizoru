import json
import re
import io
import time
import zipfile
import urllib.parse
from pathlib import Path
from typing import Any, List, Dict
from datetime import datetime, timedelta
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
        "load_batch": "📁 Ham .zip / .json (followers, following) Dosyalarını Yükleyin",
        "btn_analyze": "ANALİZİ BAŞLAT",
        "tab_unfollowers": "Beni Takip Etmeyenler",
        "tab_fans": "Geri Takip Etmediklerim",
        "tab_my_following": "Benim Takip Ettiklerim",
        "tab_ghosts": "Hayalet (Ghost) Hesaplar",
        "input_error_msg": "Analiz için veri havuzunda en az 'followers.json' ve 'following.json' bulunmalıdır.",
        "parse_error_msg": "Yüklenen JSON şeması siber motor tarafından çözümlenemedi.",
        "success_msg": "Dosyalar tarandı, Excel raporları köprü linkleriyle üretildi!",
        "perfect_sync": "🎉 [KUSURSUZ SENKRONİZASYON]: Herkes sizi geri takip ediyor!",
        "no_fans": "🎯 [HAYRAN YOK]: Takip ettiğiniz herkesi siz de geri takip ediyorsunuz.",
        "no_my_following": "📢 [TAKİP LİSTESİ BOŞ]: Takip ettiğiniz hiç kimse bulunamadı.",
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
        "chart_title": "📈 Profil Analiz Dağılımı",
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
        "logout_btn": "🔒 OTURUMU GÜVEN GİRİŞİ KAPAT (ÇIKIŞ YAP)",
        "premium_notice": "👑 **PREMIUM ÖZELLİK KİLİTLİ:** Advanced grafik analizler, Excel indirme motoru ve kronolojik zaman sıralaması sadece Premium üyelere özeldir. Yetki yükseltmek için lütfen yöneticiyle iletişime geçin.",
        "badge_premium": "👑 Premium Hesap (Sınırsız Erişim)",
        "badge_standard": "👤 Standart Hesap (Kısıtlı Erişim)",
        "welcome_user": "Hoş geldiniz, {user}"
    },
    "EN": {
        "main_title": "THREADS GLOBAL / MURAT ŞENER PRODUCER",
        "main_sub": "LOCAL AND SECURE BI-DIRECTIONAL PROFILE ANALYSIS SYSTEM // MOBILE ACT & WEB",
        "main_hashtag": "#notmebutwe",
        "load_batch": "📁 Upload Raw .zip or Multiple .json (followers, following) Files",
        "btn_analyze": "START ANALYSIS",
        "tab_unfollowers": "Not Following Me Back",
        "tab_fans": "I Am Not Following Back",
        "tab_my_following": "Whom I Follow",
        "tab_ghosts": "Ghost / Inactive Accounts",
        "input_error_msg": "Analysis requires at least 'followers.json' and 'following.json' in the pool.",
        "parse_error_msg": "The uploaded JSON schema could not be resolved by the engine.",
        "success_msg": "Files scanned, Excel reports generated with clickable links!",
        "perfect_sync": "🎉 [PERFECT SYNC]: Everyone is following you back!",
        "no_fans": "🎯 [NO FANS]: You are following back everyone who follows you.",
        "no_my_following": "📢 [FOLLOWING LIST EMPTY]: No accounts found in your following list.",
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
        "chart_title": "📈 Profile Analysis Distribution",
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
        "logout_btn": "🔒 SECURE LOG-OUT",
        "premium_notice": "👑 **PREMIUM FEATURE LOCKED:** Advanced charts, Excel exports, and chronological sorting are restricted to Premium members.",
        "badge_premium": "👑 Premium Account (Unrestricted)",
        "badge_standard": "👤 Standard Account (Restricted)",
        "welcome_user": "Welcome, {user}"
    }
}
# --- YARDIMCI GEÇİCİ KOD ÜRETME FONKSİYONU ---
def kod_uret():
    db_str = str(st.session_state.get("user_db", {}))
    prem_str = str(st.session_state.get("premium_users", set()))
    return f"""# Güncel Veritabanı Durumu\nst.session_state.user_db = {db_str}\nst.session_state.premium_users = {prem_str}"""

# --- SİBER JET HIZLANDIRICI ÖNBELLEK MOTORU ---
@st.cache_data(show_spinner=False)
def siber_json_coz(raw_bytes: bytes) -> str:
    return raw_bytes.decode("utf-8")

class AnalizMotoru:
    """Kullanıcı adlarını, zaman damgalarını, bot riskini bir kerede işleyen siber motor."""
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
        if not username: return False
        if re.search(r'\d{3,}', username): return True
        if timestamp > 0:
            fark_gun = (datetime.now() - datetime.fromtimestamp(timestamp)).days
            if fark_gun > 540: return True
        return False

    # --- 🚨 YENİ: ÇOKLU TOPLU LİNK AÇICI SİBER JAVASCRIPT MOTORU ---
    @staticmethod
    def toplu_link_ac(kullanici_listesi: List[str], limit: int):
        if not kullanici_listesi: return
        hedef_liste = kullanici_listesi[:limit]
        js_kodları = ""
        for u in hedef_liste:
            js_kodları += f"window.open('https://threads.com/@{u}', '_blank');\n"
        
        st.components.v1.html(f"<script>{js_kodları}</script>", height=0, width=0)
# --- 🔑 GÜVENLİ VE KALICI DEĞİŞTİRİLEBİLİR KOD TABANLI VERİ TABANI ---
if "user_db" not in st.session_state:
    st.session_state.user_db = {"murat": "snr", "demo": "demo", "alkan": "alkan", "büşra": "büşra", "azat": "azat","ali": "aksoy",}

if "premium_users" not in st.session_state:
    st.session_state.premium_users = {"murat", "alkan", "büşra", "azat"}

if "islem_yapilanlar" not in st.session_state:
    st.session_state.islem_yapilanlar = set()

# --- 🚨 YENİ: TAKİPTEN ÇIKANLAR CANLI GEÇMİŞ HAFIZASI ---
if "unfollowed_history" not in st.session_state:
    st.session_state.unfollowed_history = set()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.markdown(f"### 🎯 Threads Profil Takip Sistemi")
    with st.container(border=True):
        st.markdown("<h4 style='color:#3a7ebf; margin-top:0px;'><b>🔒 SİSTEME GÜVENLİ GİRİŞ</b></h4>", unsafe_allow_html=True)
        input_user = st.text_input("👤 Kullanıcı Adı", key="login_username_field").strip().lower()
        input_pass = st.text_input("🔑 Şifre", type="password", key="login_password_field").strip()
        if st.button("SİSTEME GİRİŞ YAP", use_container_width=True, type="primary"):
            if input_user in st.session_state.user_db and st.session_state.user_db[input_user] == input_pass:
                st.session_state.logged_in = True
                st.session_state.current_active_user = input_user
                st.success("🔓 Erişim Onaylandı! Sistem yükleniyor...")
                st.rerun()
            else: st.error("❌ Hatalı Kullanıcı Adı veya Şifre!")
    st.stop()

# --- SESSİON_STATE TABANLI TEMA HAFIZASI ---
if "sabit_tema" not in st.session_state: st.session_state.sabit_tema = "Karanlık Gece Modu"
if st.session_state.sabit_tema == "Premium Fildişi & Kemik Modu":
    st.markdown("<style>.stApp { background-color: #f9f6f0 !important; color: #1c1c1e !important; } h1, h2, h3, h4, h5, h6, p, label, span, small { color: #1c1c1e !important; } div[data-testid='stExpander'], div[data-testid='stFileUploader'], div[data-testid='stDataframe'] { background-color: #f1ede4 !important; border: 1px solid #e1dacb !important; } .stMarkdown p { color: #1c1c1e !important; } .stMarkdown b { color: #000000 !important; }</style>", unsafe_allow_html=True)
st.title("🎯 Threads Profil Takip Sistemi")
aktif_u = st.session_state.current_active_user
is_user_premium = (aktif_u in st.session_state.premium_users)

col_user_welcome, col_user_badge = st.columns(2)
with col_user_welcome: st.markdown(f"👋 **{DIL_PAKETI['TR']['welcome_user'].format(user=aktif_u)}**")
with col_user_badge:
    if is_user_premium: st.markdown(f"<p style='text-align: right; margin: 0; font-weight: bold; color: #2e7d32;'>{DIL_PAKETI['TR']['badge_premium']}</p>", unsafe_allow_html=True)
    else: st.markdown(f"<p style='text-align: right; margin: 0; font-weight: bold; color: #ef6c00;'>{DIL_PAKETI['TR']['badge_standard']}</p>", unsafe_allow_html=True)

col_lang, col_hashtag = st.columns(2)
with col_lang: aktif_dil = st.selectbox("🌐 Language / Dil", ["TR", "EN"])
with col_hashtag: st.markdown(f"<h4 style='text-align: right; color: #3a7ebf; margin-top: 5px;'>{DIL_PAKETI[aktif_dil]['main_hashtag']}</h4>", unsafe_allow_html=True)

st.markdown(f"### {DIL_PAKETI[aktif_dil]['main_title']}"); st.caption(DIL_PAKETI[aktif_dil]['main_sub']); st.divider()
st.link_button(label="▶️ YAPARKEN DİNLERSİNİZ BELKİ (Göndermeli Şarkı)", url="https://www.youtube.com/watch?v=7S-E0spllUM", use_container_width=True)

with st.expander(DIL_PAKETI[aktif_dil]['guide_title'], expanded=False):
    st.info(DIL_PAKETI[aktif_dil]['pwa_guide_text'])
    for s in ['guide_step1', 'guide_step2', 'guide_step3', 'guide_step4', 'guide_step5', 'guide_step6']: st.markdown(DIL_PAKETI[aktif_dil][s])

uploaded_inputs = st.file_uploader(DIL_PAKETI[aktif_dil]['load_batch'], type=["zip", "json"], accept_multiple_files=True)
following_bytes, followers_bytes = None, None

if uploaded_inputs:
    for item in uploaded_inputs:
        if item.name.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(item) as z:
                    for f_info in z.infolist():
                        fname = f_info.filename.lower()
                        if "following.json" in fname: following_bytes = z.read(f_info.filename)
                        elif "followers.json" in fname: followers_bytes = z.read(f_info.filename)
            except Exception: st.error("Zip Çözümleme Hatası!")
        elif item.name.lower().endswith(".json"):
            fname = item.name.lower()
            if "following" in fname: following_bytes = item.read()
            elif "followers" in fname: followers_bytes = item.read()
if st.button(DIL_PAKETI[aktif_dil]['btn_analyze'], use_container_width=True, type="primary") or st.session_state.get('analyzed', False):
    if not following_bytes or not followers_bytes: st.warning(DIL_PAKETI[aktif_dil]['input_error_msg'])
    else:
        try:
            if 'unfollowers' not in st.session_state:
                # Hafıza karşılaştırması için eski takipçi kümesini koruma altına alma
                eski_takipciler = st.session_state.get("followers_set", set())

                st.session_state.global_following_map = AnalizMotoru.akilli_süre_ayristir(json.loads(siber_json_coz(following_bytes)))
                st.session_state.global_followers_map = AnalizMotoru.akilli_süre_ayristir(json.loads(siber_json_coz(followers_bytes)))
                st.session_state.following_set, st.session_state.followers_set = set(st.session_state.global_following_map.keys()), set(st.session_state.global_followers_map.keys())
                
                # --- 🚨 YENİ: UNFOLLOW (TAKİPTEN ÇIKANLAR) TESPİT ALGORİTMASI ---
                if len(eski_takipciler) > 0:
                    yeni_unfollowers = eski_takipciler - st.session_state.followers_set
                    if yeni_unfollowers:
                        st.session_state.unfollowed_history.update(yeni_unfollowers)

                st.session_state.current_unfollowers = st.session_state.following_set - st.session_state.followers_set
                st.session_state.current_fans = st.session_state.followers_set - st.session_state.following_set
                st.session_state.current_my_following = st.session_state.following_set
                st.session_state.ghosts = {u for u, ts in st.session_state.global_followers_map.items() if AnalizMotoru.bot_ve_pasiflik_kontrolü(u, ts)}
                st.session_state.analyzed = True

            global_following_map, global_followers_map = st.session_state.global_following_map, st.session_state.global_followers_map
            following_set, followers_set = st.session_state.following_set, st.session_state.followers_set
            unfollowers, fans, my_following, ghosts = st.session_state.current_unfollowers, st.session_state.current_fans, st.session_state.current_my_following, st.session_state.ghosts

            if not following_set or not followers_set: st.error(DIL_PAKETI[aktif_dil]['parse_error_msg'])
            else:
                if st.session_state.get("chk_outdated_alert", True):
                    en_son_sinyal = max(list(global_following_map.values()) + list(global_followers_map.values()), default=0)
                    if en_son_sinyal > 0:
                        gecen_gun = (datetime.now() - datetime.fromtimestamp(en_son_sinyal)).days
                        if gecen_gun > 30: st.warning(DIL_PAKETI[aktif_dil]['outdated_warning'].replace("{days}", str(gecen_gun)))

                islem_yapilanlar = st.session_state.islem_yapilanlar
                dinamik_unfollowers = unfollowers - islem_yapilanlar
                dinamik_fans = fans - islem_yapilanlar
                dinamik_my_following = my_following - islem_yapilanlar
                dinamik_ghosts = ghosts - islem_yapilanlar

                aktif_following_count = len(dinamik_my_following)
                aktif_followers_count = len(followers_set - (islem_yapilanlar & followers_set))

                ceza_puanı = (len(dinamik_unfollowers) / aktif_following_count) * 40 if aktif_following_count > 0 else 0
                ghost_ceza = (len(dinamik_ghosts) / aktif_followers_count) * 20 if aktif_followers_count > 0 else 0
                denge_puanı = (min(aktif_followers_count, aktif_following_count) / max(aktif_followers_count, aktif_following_count)) * 40 if aktif_following_count and aktif_followers_count else 0
                health_score = max(0, min(100, int(100 - ceza_puanı - ghost_ceza + (denge_puanı * 0.1))))
                
                durum_str = "MÜKEMMEL" if health_score > 85 else "STABİL" if health_score > 60 else "RİSKLİ"
                st.success(DIL_PAKETI[aktif_dil]['success_msg']); st.subheader(DIL_PAKETI[aktif_dil]['summary_title'])
                
                m1, m2, m3 = st.columns(3)
                m1.metric(DIL_PAKETI[aktif_dil]['health_score'], f"%{health_score}", durum_str)
                m2.metric("Following (Net)", aktif_following_count)
                m3.metric("Followers (Net)", aktif_followers_count)

                card_bg = "#121212" if st.session_state.sabit_tema == "Karanlık Gece Modu" else "#f1ede4"
                card_text = "#ffffff" if st.session_state.sabit_tema == "Karanlık Gece Modu" else "#1c1c1e"
                st.markdown(f"<div style='background-color:{card_bg}; border:2px solid #3a7ebf; border-radius:15px; padding:25px; margin-top:15px; text-align:center; font-family:-apple-system,BlinkMacSystemFont;'><h3 style='color:#3a7ebf; margin-bottom:5px; font-weight:bold;'>🎯 THREADS DURUM</h3><p style='color:{card_text}; font-size:14px;'>@ {st.session_state.current_active_user} &nbsp;|&nbsp; {datetime.now().strftime('%Y-%m-%d')}</p><hr style='border:0; border-top:1px solid #3a7ebf; margin:15px 0;'><div style='display:flex; justify-content:space-around; margin:20px 0;'><div><span style='font-size:12px; color:#888;'>SAĞLIK SKORU</span><br><b style='font-size:26px; color:#3a7ebf;'>%{health_score}</b></div><div><span style='font-size:12px; color:#888;'>TAKİPÇİLER</span><br><b style='font-size:26px; color:{card_text};'>{aktif_followers_count}</b></div><div><span style='font-size:12px; color:#888;'>TAKİP ETTİKLERİN</span><br><b style='font-size:26px; color:{card_text};'>{aktif_following_count}</b></div></div></div>", unsafe_allow_html=True)

                if st.session_state.current_active_user in st.session_state.premium_users:
                    st.write(""); st.markdown(f"##### {DIL_PAKETI[aktif_dil]['chart_title']}")
                    st.bar_chart(data={"Kategori": ["Takip Etmeyenler", "Karşılıklı", "Hayranlar", "Takip Ettiklerim", "Hayaletler"], "Sayı": [len(dinamik_unfollowers), len(dinamik_my_following & (followers_set - islem_yapilanlar)), len(dinamik_fans), len(dinamik_my_following), len(dinamik_ghosts)]}, x="Kategori", y="Sayı", use_container_width=True)
                st.info(DIL_PAKETI[aktif_dil]['premium_notice'])
                output_excel = io.BytesIO()
                workbook = xlsxwriter.Workbook(output_excel)
                header_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#1f4e78', 'border': 1, 'align': 'center'})

                sheet_unf = workbook.add_worksheet("Beni Takip Etmeyenler")
                sheet_unf.write_row('A1', ['No', 'Kullanıcı Adı', 'Profil Linki', 'Süre'], header_format)
                for idx, user in enumerate(sorted(dinamik_unfollowers, key=lambda x: global_following_map.get(x, 0)), 1):
                    sheet_unf.write_row(idx, 0, [idx, f"@{user}", f"https://threads.com/@{user}", AnalizMotoru.zaman_metnine_cevir(global_following_map.get(user, 0))])
                sheet_unf.set_column('B:C', 25); sheet_unf.set_column('C:C', 45)

                sheet_fans = workbook.add_worksheet("Geri Takip Etmediklerim")
                sheet_fans.write_row('A1', ['No', 'Kullanıcı Adı', 'Profil Linki', 'Süre'], header_format)
                for idx, user in enumerate(sorted(dinamik_fans, key=lambda x: global_followers_map.get(x, 0)), 1):
                    sheet_fans.write_row(idx, 0, [idx, f"@{user}", f"https://threads.com/@{user}", AnalizMotoru.zaman_metnine_cevir(global_followers_map.get(user, 0))])
                sheet_fans.set_column('B:C', 25); sheet_fans.set_column('C:C', 45)

                sheet_my_f = workbook.add_worksheet("Benim Takip Ettiklerim")
                sheet_my_f.write_row('A1', ['No', 'Kullanıcı Adı', 'Profil Linki', 'Süre'], header_format)
                for idx, user in enumerate(sorted(dinamik_my_following, key=lambda x: global_following_map.get(x, 0)), 1):
                    sheet_my_f.write_row(idx, 0, [idx, f"@{user}", f"https://threads.com/@{user}", AnalizMotoru.zaman_metnine_cevir(global_following_map.get(user, 0))])
                sheet_my_f.set_column('B:C', 25); sheet_my_f.set_column('C:C', 45)

                sheet_gh = workbook.add_worksheet("Hayalet Hesaplar")
                sheet_gh.write_row('A1', ['No', 'Kullanıcı Adı', 'Profil Linki', 'Süre'], header_format)
                for idx, user in enumerate(sorted(dinamik_ghosts, key=lambda x: global_followers_map.get(x, 0)), 1):
                    sheet_gh.write_row(idx, 0, [idx, f"@{user}", f"https://threads.com/@{user}", AnalizMotoru.zaman_metnine_cevir(global_followers_map.get(user, 0))])
                sheet_gh.set_column('B:C', 25); sheet_gh.set_column('C:C', 45); workbook.close(); output_excel.seek(0)
                
                if st.session_state.current_active_user in st.session_state.premium_users:
                    st.download_button(label=DIL_PAKETI[aktif_dil]['download_excel'], data=output_excel, file_name="Threads_Detayli_Analiz_Raporu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                is_reverse = (st.radio(label="", options=[DIL_PAKETI[aktif_dil]['sort_newest'], DIL_PAKETI[aktif_dil]['sort_oldest']], horizontal=True, label_visibility="collapsed") == DIL_PAKETI[aktif_dil]['sort_newest']) if st.session_state.current_active_user in st.session_state.premium_users else True
                clean_query = st.text_input("", placeholder=DIL_PAKETI[aktif_dil]['search_placeholder']).strip().lower().replace("@", "")

                sorted_unf = sorted(dinamik_unfollowers, key=lambda x: global_following_map.get(x, 0), reverse=is_reverse)
                sorted_fans = sorted(dinamik_fans, key=lambda x: global_followers_map.get(x, 0), reverse=is_reverse)
                sorted_my_following = sorted(dinamik_my_following, key=lambda x: global_following_map.get(x, 0), reverse=is_reverse)
                sorted_gh = sorted(dinamik_ghosts, key=lambda x: global_followers_map.get(x, 0), reverse=is_reverse)

                t1, t2, t3, t4 = st.tabs([DIL_PAKETI[aktif_dil]['tab_unfollowers'], DIL_PAKETI[aktif_dil]['tab_fans'], DIL_PAKETI[aktif_dil]['tab_my_following'], DIL_PAKETI[aktif_dil]['tab_ghosts']])
                
                def render_list(target_list, prefix, is_following_map=True):
                    fil = [u for u in target_list if clean_query in u.lower()] if clean_query else target_list
                    
                    # --- 🚨 YENİ: ÇOKLU TOPLU LİNK AÇICI MOBİL BUTONLARI ENJEKTE EDİLDİ ---
                    if fil:
                        lc1, lc2, _ = st.columns([0.35, 0.35, 0.30])
                        with lc1:
                            if st.button("🚀 İlk 5 Profili Aç", key=f"toplu_5_{prefix}", use_container_width=True):
                                AnalizMotoru.toplu_link_ac(fil, 5)
                        with lc2:
                            if st.button("🚀 İlk 10 Profili Aç", key=f"toplu_10_{prefix}", use_container_width=True):
                                AnalizMotoru.toplu_link_ac(fil, 10)
                        st.write("")

                        for index, user in enumerate(fil, 1):
                            c_item, c_btn = st.columns([0.75, 0.25])
                            with c_item:
                                ts_v = global_following_map.get(user, 0) if is_following_map else global_followers_map.get(user, 0)
                                st.markdown(f"[{index:03d}] 🔗 [@{user}](https://threads.com/@{user}) &nbsp;&nbsp;&nbsp;&nbsp; <b>⌛ {AnalizMotoru.zaman_metnine_cevir(ts_v)}</b>", unsafe_allow_html=True)
                            with c_btn:
                                if st.button("Listeden Kaldır", key=f"btn_done_{prefix}_{user}_{index}"):
                                    st.session_state.islem_yapilanlar.add(user)
                                    st.rerun()
                    else: st.info("Gösterilecek hesap kalmadı.")

                with t1: render_list(sorted_unf, "unf", is_following_map=True)
                with t2: render_list(sorted_fans, "fans", is_following_map=False)
                with t3: render_list(sorted_my_following, "myf", is_following_map=True)
                with t4: render_list(sorted_gh, "ghost", is_following_map=False)
        except Exception as e: st.error(f"Sistem Hatası: {str(e)}")

st.write(""); st.divider()

# --- 📋 1. KATALOG: İŞLEM YAPILANLARIN LİSTESİ ---
with st.expander(f"📋 İşlem Yapılanların Listesi ({len(st.session_state.islem_yapilanlar)})", expanded=False):
    if len(st.session_state.islem_yapilanlar) > 0:
        st.markdown("<p style='color:#3a7ebf; font-weight:bold;'>✔️ İŞLEM YAPILAN HESAPLAR</p>", unsafe_allow_html=True)
        for i_idx, i_user in enumerate(sorted(list(st.session_state.islem_yapilanlar)), 1):
            col_u_name, col_undo = st.columns([0.75, 0.25])
            with col_u_name: st.markdown(f"[{i_idx:03d}] 🔗 [@{i_user}](https://threads.com/@{i_user})", unsafe_allow_html=True)
            with col_undo:
                if st.button("↩️ Geri Al", key=f"undo_{i_user}_{i_idx}"):
                    st.session_state.islem_yapilanlar.discard(i_user)
                    st.rerun()
        st.divider()
    if st.button("🔄 Tüm İşlem Yapılan Listesini Sıfırla", use_container_width=True):
        st.session_state.islem_yapilanlar.clear(); st.success("Tüm hesaplar listelere geri yüklendi!"); st.rerun()

# --- 🚨 2. KATALOG: YENİ EKLENEN TAKİPTEN ÇIKANLAR CANLI GEÇMİŞİ ---
with st.expander(f"🚨 Takipten Çıkanlar Canlı Geçmişi ({len(st.session_state.unfollowed_history)})", expanded=False):
    if len(st.session_state.unfollowed_history) > 0:
        st.markdown("<p style='color:#d32f2f; font-weight:bold;'>⚠️ SİZİ YAKIN ZAMANDA TAKİPTEN ÇIKARANLAR</p>", unsafe_allow_html=True)
        for h_idx, h_user in enumerate(sorted(list(st.session_state.unfollowed_history)), 1):
            st.markdown(f"[{h_idx:03d}] 🚨 [@{h_user}](https://threads.com/@{h_user})", unsafe_allow_html=True)
    else:
        st.info("Bu oturumda henüz sizi takipten çıkaran bir hesap algılanmadı. Yeni dosya yüklediğinizde geçmiş hafızası tetiklenecektir.")

if is_user_premium and aktif_u == "murat":
    with st.expander("👑 SaaS Üye Paneli (Admin)", expanded=False):
        s_üye = st.selectbox("Üye Seçin", list(st.session_state.user_db.keys()))
        if s_üye:
            g_pass = st.text_input(f"New Password", type="password", key="saas_edit_pass")
            g_role = st.radio(f"Role", ["Standart (Kısıtlı)", "Premium (Limitsiz)"], index=1 if s_üye in st.session_state.premium_users else 0, horizontal=True)
            if st.button("ÜYE GÜNCELLE", use_container_width=True, type="primary"):
                if g_pass.strip(): st.session_state.user_db[s_üye] = g_pass.strip()
                st.session_state.premium_users.add(s_üye) if g_role == "Premium (Limitsiz)" else st.session_state.premium_users.discard(s_üye)
                st.success("Güncellendi!"); st.rerun()
        st.divider(); st.code(kod_uret(), language="python")

with st.expander("⚙️ Gelişmiş Hesap Ayarları", expanded=False):
    st.write(f"• **User:** `@{aktif_u}` | **Role:** `{'👑 Premium' if is_user_premium else '👤 Standart'}`")
    if st.radio("Sistem Teması", ["Karanlık Gece Modu", "Premium Fildişi & Kemik Modu"], index=0 if st.session_state.sabit_tema == "Karanlık Gece Modu" else 1) != st.session_state.sabit_tema:
        st.session_state.sabit_tema = "Karanlık Gece Modu" if st.session_state.sabit_tema != "Karanlık Gece Modu" else "Premium Fildişi & Kemik Modu"
        st.success("Tema Değişti!"); st.rerun()
    st.divider(); st.session_state.chk_outdated_alert = st.checkbox("Eski Veri Uyarısı", value=st.session_state.get("chk_outdated_alert", True))
    st.divider(); yeni_sifre = st.text_input("Yeni Şifre Girişi", type="password", key="change_password_box").strip()
    if st.button("ŞİFREYİ GÜNCELLE (muratsenr ile iletişime geçiniz)", use_container_width=True) and yeni_sifre:
        st.session_state.user_db[aktif_u] = yeni_sifre; st.success("Şifre güncellendi!"); st.code(kod_uret(), language="python")

st.write("")
if st.button("🗑️ OTURUMU KAPAT VE TÜM VERİLERİ TEMİZLE (DEEP CLEAN)", use_container_width=True, type="secondary"):
    st.cache_data.clear()
    for k in list(st.session_state.keys()): del st.session_state[k]
    st.session_state.logged_in, st.session_state.analyzed = False, False; st.rerun()
st.write(""); st.link_button(label=DIL_PAKETI[aktif_dil]['contact_btn'], url="https://threads.com/@muratsenr", use_container_width=True)
