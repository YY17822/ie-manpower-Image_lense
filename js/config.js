/*
 * config.js — 這個專案的身分設定，網頁和 sw.js（離線快取）都讀這一個檔案。
 *
 * 複製這個專案做成新專案時，先改這裡的 PROJECT_ID（每個專案都要不同），
 * 瀏覽器裡的快取和記住的設定才會和原本的專案分開，不會互相讀到、蓋掉或刪掉。
 *   快取名稱       PROJECT_ID + '-v' + 版本號，例如 ie-manpower-3F-image-lens-v26（版本號在 sw.js）
 *   記住的設定     PROJECT_ID + '.' + 名稱，例如 ie-manpower-3F-image-lens.step1
 *
 * SUPABASE_*：共用設定的資料庫，複製成新專案時不用改（同一個 Supabase、同一個 schema），只要 PROJECT_ID 不同就不會混在一起。
 *   換資料庫（雲端 → 本機自架 → 公司主機）時只改 SUPABASE_URL 和 SUPABASE_KEY，程式不用動：
 *     Supabase 雲端         https://xxxx.supabase.co               Settings → API Keys 的 publishable key
 *     Mac 上 supabase CLI   http://127.0.0.1:54321                 `supabase status` 顯示的 Publishable（或 anon）key
 *     Docker 自架           http://<主機 IP 或名稱>:8000            docker 資料夾 .env 裡的 ANON_KEY
 *   只能放 publishable／anon key，secret／service_role key 絕對不能放進來。
 *
 * LEGACY_*：這個專案以前用的舊名稱，新版第一次打開時會把舊設定搬過來、把舊快取刪掉。
 *           新專案用不到，複製之後請改成 null。
 */
(function (g) {
  'use strict';
  g.APP_CONFIG = {
    PROJECT_ID: 'ie-manpower-3F-image-lens',
    LEGACY_STORE_PREFIX: 'ie-manpower.',          // 舊的設定名稱開頭
    LEGACY_CACHE_PATTERN: '^ie-manpower-v\\d+$',  // 舊的快取名稱（ie-manpower-v1、v2…）

    // 共用設定的資料庫（Supabase）。8 個人力需求 PWA 共用同一個專案和 schema，用 PROJECT_ID 分開。
    // publishable key 本來就是公開用的，可以放在網頁裡；安全靠登入和資料表的權限規則（RLS）。
    // 不要把 secret / service_role key 放進這裡。
    // SUPABASE_URL 留空字串 = 不用 Supabase，設定只存在這台電腦（和以前一樣）。
    // 雲端：SUPABASE_URL 'https://ojjplngsnovqeyuvagzi.supabase.co'、SUPABASE_KEY 'sb_publishable_aghJ5ZyZnxphMiVri6zZcQ_Fm92pmMq'
    // 本機 Docker 自架（docker-lab/supabase-project，key 是 .env 的 ANON_KEY）：
    SUPABASE_URL: 'http://localhost:8000',
    SUPABASE_KEY: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJyb2xlIjoiYW5vbiIsImlzcyI6InN1cGFiYXNlIiwiaWF0IjoxNzkyNzIxNDM4LCJleHAiOjE5NTA0MDE0Mzh9.xAa4TkR4X2nrDaZAGfgP5H6gU0eZ3kKNTRZg0WU_AqY',
    SUPABASE_SCHEMA: 'ie_manpower'
  };
})(typeof self !== 'undefined' ? self : this);
