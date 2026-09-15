/**
 * BACKEND CMS UP2DIGITAL (GOOGLE APPS SCRIPT)
 * Versi: 3.4 (Full Complete Rebuild - All Setup Functions Included)
 */
const SCRIPT_NAME = "Up2Digital_CMS_Backend_v4";
const ADMIN_PIN = "1234"; // PIN rahasia dipindah ke Backend agar aman
// ==========================================
// PENTING: GANTI DENGAN ID SPREADSHEET ANDA!
// ==========================================
const SHEET_ID = "1GPDJQ1h_rxYHW_wXbcQSgUTFrFzffb1ljx1YwpZFiHU"; 

// ==========================================
// 1. HANDLE HTTP GET (Read Data)
// ==========================================
function doGet(e) {
  let response = { status: "error", message: "Action not found" };

  try {
    if (!e || !e.parameter || !e.parameter.action) {
       return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "API Backend CMS v3.4 Aktif!" })).setMimeType(ContentService.MimeType.JSON);
    }

    const action = e.parameter.action;
    const validActions = ["getNews", "getContacts", "getCareers", "getProducts", "getAssets", "getCopywriting"];
    
    if (validActions.includes(action)) {
      const sheetName = action.replace("get", "");
      response = { status: "success", data: getSheetData(sheetName) };
    } else {
      response = { status: "error", message: "Invalid GET action specified." };
    }
  } catch(err) {
    response = { status: "error", message: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 2. HANDLE HTTP POST (Create / Update / Delete)
// ==========================================
function doPost(e) {
  let response = { status: "error", message: "Invalid request" };
  
  try {
    if (!e || !e.postData || !e.postData.contents) {
        throw new Error("No data received in POST request.");
    }

    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const payload = data.payload;

    if (!action) {
         throw new Error("No action specified in POST request.");
    }

    if (action === "verifyPin") {
      if (payload.pin === ADMIN_PIN) {
        response = { status: "success", message: "Login berhasil" };
      } else {
        response = { status: "error", message: "PIN Salah" };
      }
    }
    else if (action === "submitContact") {
      response = appendRowData("Contacts", [new Date().toISOString(), payload.name, payload.email, payload.phone, payload.topic, payload.message]);
    }
    else if (action === "addCareers") {
      const id = "job_" + new Date().getTime().toString();
      response = appendRowData("Careers", [id, payload.title, payload.department, payload.location, payload.status, new Date().toISOString()]);
    }
    else if (action === "addProducts") {
      const id = payload.id || "prod_" + new Date().getTime().toString();
      response = appendRowData("Products", [id, payload.product_name, payload.description, payload.link_url, payload.link_text, payload.image_url || ""]);
    }
    else if (action.startsWith("delete")) {
       const sheetName = action.replace("delete", ""); 
       response = deleteRowById(sheetName, payload.id);
    }
    else if (action.startsWith("update")) {
       const sheetName = action.replace("update", ""); 
       response = updateRowById(sheetName, payload.id, payload);
    } else {
       throw new Error("Invalid POST action specified: " + action);
    }
  } catch (err) {
    response = { status: "error", message: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// 3. HELPER FUNCTIONS 
// ==========================================
function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet " + sheetName + " tidak ditemukan.");
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; 
  
  const headers = data.shift();
  const rows = data.map(row => {
    let obj = {};
    headers.forEach((header, i) => { obj[header] = row[i]; });
    return obj;
  });
  
  if(["News", "Contacts", "Careers"].includes(sheetName)) return rows.reverse();
  return rows;
}

function appendRowData(sheetName, rowData) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet " + sheetName + " tidak ditemukan.");
  sheet.appendRow(rowData);
  return { status: "success", message: "Data berhasil disimpan di " + sheetName };
}

function deleteRowById(sheetName, id) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet " + sheetName + " tidak ditemukan.");
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { 
      sheet.deleteRow(i + 1); 
      return { status: "success", message: "Data berhasil dihapus." };
    }
  }
  throw new Error("Data ID " + id + " tidak ditemukan.");
}

function updateRowById(sheetName, id, payload) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet " + sheetName + " tidak ditemukan.");
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { 
      let updatedRow = [...data[i]]; 
      headers.forEach((header, colIndex) => {
        if (payload[header] !== undefined) {
          updatedRow[colIndex] = payload[header];
        }
      });
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([updatedRow]);
      return { status: "success", message: "Data berhasil diupdate." };
    }
  }
  throw new Error("Data ID " + id + " tidak ditemukan.");
}

// ==========================================
// 4. SETUP ENVIRONMENT (RUN ONCE)
// ==========================================
function setupEnvironment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ss.getSheetByName("News")) { ss.insertSheet("News").appendRow(["id", "title", "date", "category", "author", "content", "thumbnail"]); }
  if (!ss.getSheetByName("Contacts")) { ss.insertSheet("Contacts").appendRow(["timestamp", "name", "email", "phone", "topic", "message"]); }
  if (!ss.getSheetByName("Careers")) { ss.insertSheet("Careers").appendRow(["id", "title", "department", "location", "status", "created_at"]); }

  let prodSheet = ss.getSheetByName("Products");
  if (!prodSheet) {
    prodSheet = ss.insertSheet("Products");
    prodSheet.appendRow(["id", "product_name", "description", "link_url", "link_text", "image_url"]);
    prodSheet.appendRow(["prod_gasora", "GASORA", "Infrastruktur aplikasi bisnis cerdas...", "https://gasora.up2digital.workers.dev/", "Selengkapnya", ""]);
    prodSheet.appendRow(["prod_webkit", "WebKit", "Solusi website & landing page...", "https://webkit.up2digital.workers.dev/", "Selengkapnya", ""]);
    prodSheet.appendRow(["prod_blanjaan", "Blanjaan", "Digital marketplace...", "#", "Coming Soon", ""]);
  }

  let copySheet = ss.getSheetByName("Copywriting");
  if (!copySheet) {
    copySheet = ss.insertSheet("Copywriting");
    copySheet.appendRow(["id", "section_name", "text_content"]);
    copySheet.appendRow(["copy_hero_title", "Hero Title", "Membangun Ekosistem Digital untuk Indonesia"]);
    copySheet.appendRow(["copy_founder_quote", "Founder Quote", "Teknologi Seharusnya Memberdayakan, Bukan Membingungkan."]);
    copySheet.appendRow(["copy_link_ig", "Instagram Link", "https://www.instagram.com/up2.media"]);
    copySheet.appendRow(["copy_link_yt", "Youtube Link", "https://www.youtube.com/@JajamanOfficial"]);
    copySheet.appendRow(["copy_link_tiktok", "TikTok Link", "https://www.tiktok.com/@up2.media"]);
    copySheet.appendRow(["copy_link_fb", "Facebook Link", "https://www.facebook.com"]);
    copySheet.appendRow(["copy_link_threads", "Threads Link", "https://www.threads.com/@up2.media"]);
    copySheet.appendRow(["copy_link_blogger", "Blogger Link", "https://kulineriadigital.blogspot.com/"]);
  }
  
  Logger.log("==== SETUP DATABASE SELESAI ====");
}

// ==========================================
// 5. FORCE UPDATE ASSETS (FUNGSI PERBAIKAN)
// ==========================================
function forceUpdateAssets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let assetSheet = ss.getSheetByName("Assets");
  
  if (!assetSheet) { assetSheet = ss.insertSheet("Assets"); }
  
  assetSheet.clear();
  assetSheet.appendRow(["id", "element_name", "asset_url", "type"]);
  
  assetSheet.appendRow(["asset_favicon", "Favicon", "https://raw.githubusercontent.com/up2digi-app/up2digital/main/favicon.ico", "Image"]);
  assetSheet.appendRow(["asset_logo_color_1", "Logo Utama Berwarna (1)", "https://raw.githubusercontent.com/up2digi-app/up2digital/main/logo-colour-nbg.png", "Image"]);
  assetSheet.appendRow(["asset_logo_color_2", "Logo Utama Berwarna (2)", "https://raw.githubusercontent.com/up2digi-app/up2digital/main/logo-colour-nbg.png", "Image"]);
  assetSheet.appendRow(["asset_logo_white_1", "Logo Putih (1)", "https://raw.githubusercontent.com/up2digi-app/up2digital/main/logo-white.png", "Image"]);
  assetSheet.appendRow(["asset_logo_white_2", "Logo Putih (2)", "https://raw.githubusercontent.com/up2digi-app/up2digital/main/logo-white.png", "Image"]);
  
  assetSheet.appendRow(["asset_hero_img", "Hero Utama Beranda", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/Entrepreneurs_collaborating.jpg", "Image"]);
  assetSheet.appendRow(["asset_about_img", "Gambar Tentang Kami", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/support-umkm.jpg", "Image"]);
  assetSheet.appendRow(["asset_founder", "Foto Utama Pesan Founder (.jpeg)", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/founder-up2digital.jpeg", "Image"]);
  assetSheet.appendRow(["asset_commit_bg", "Background Komitmen", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/ilustrasi_komitmen-kami.jpg", "Image"]);
  assetSheet.appendRow(["asset_node_apps", "Node UP2Apps", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/ilustrasi_up2app.jpg", "Image"]);
  assetSheet.appendRow(["asset_node_commerce", "Node UP2Commerce", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/ilustrasi_up2commerce.jpg", "Image"]);
  assetSheet.appendRow(["asset_node_media", "Node UP2Media", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/ilustrasi_up2media.jpg", "Image"]);
  assetSheet.appendRow(["asset_node_studio", "Node UP2Studio", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/ilustrasi_up2studio.jpg", "Image"]);
  assetSheet.appendRow(["asset_node_edu", "Node UP2Edu", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/ilustrasi_up2edu.jpg", "Image"]);
  assetSheet.appendRow(["asset_prod_gasora", "Split Panel GASORA", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/landingpage-gasora.jpg", "Image"]);
  assetSheet.appendRow(["asset_prod_webkit", "Split Panel WebKit", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/landingpage-webkit.jpg", "Image"]);
  assetSheet.appendRow(["asset_prod_blanjaan", "Split Panel Blanjaan", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/landingpage-blanjaan.jpg", "Image"]);
  assetSheet.appendRow(["asset_prod_heritaku", "Split Panel HeritaKu", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/landingpage-heritaku.jpg", "Image"]);
  assetSheet.appendRow(["asset_prod_designai", "Split Panel UP2Design AI", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/landingpage-designai.jpg", "Image"]);
  assetSheet.appendRow(["asset_news_hero", "Hero Ruang Berita & Media", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/heropage_ruang-berita-01.jpg", "Image"]);
  assetSheet.appendRow(["asset_career_hero", "Hero Lowongan Karier", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/heropage_karir.jpg", "Image"]);
  assetSheet.appendRow(["asset_contact_img", "Background Ilustrasi Form Kontak", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/heropage_hubungi-kami-01.jpg", "Image"]);

  assetSheet.appendRow(["asset_hero_vid", "Video Latar Ekosistem", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/Wireframe-globe-emitting-laser.mp4", "Video"]);
  assetSheet.appendRow(["asset_contact_vid", "Video Latar Header Hubungi Kami", "https://raw.githubusercontent.com/up2digi-app/up2digital-assets/main/background/heropage-hubungi-kami.mp4", "Video"]);
  
  Logger.log("==== 24 LINK ASET BERHASIL DIPERBARUI SECARA PAKSA ====");
}

// ==========================================
// 6. PATCH DATABASE (FUNGSI PERBAIKAN)
// ==========================================
function patchDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let prodSheet = ss.getSheetByName("Products");
  if (prodSheet) {
     const data = prodSheet.getDataRange().getValues();
     const ids = data.map(r => r[0]);
     const headers = data[0];
     
     // Otomatis tambah kolom image_url di Spreadsheet jika belum ada
     if (!headers.includes("image_url")) {
       prodSheet.getRange(1, headers.length + 1).setValue("image_url");
     }

     if (!ids.includes("prod_heritaku")) prodSheet.appendRow(["prod_heritaku", "HeritaKu", "Preserve stories. Build family history...", "#", "Coming Soon", ""]);
     if (!ids.includes("prod_up2designai")) prodSheet.appendRow(["prod_up2designai", "UP2Design AI", "Solusi desain berbasis kecerdasan buatan...", "#", "Coming Soon", ""]);
  }
  
  let copySheet = ss.getSheetByName("Copywriting");
  if (copySheet) {
     const data = copySheet.getDataRange().getValues();
     const ids = data.map(r => r[0]);
     if(!ids.includes("copy_link_ig")) copySheet.appendRow(["copy_link_ig", "Instagram Link", "https://www.instagram.com/up2.media"]);
     if(!ids.includes("copy_link_yt")) copySheet.appendRow(["copy_link_yt", "Youtube Link", "https://www.youtube.com/@JajamanOfficial"]);
     if(!ids.includes("copy_link_tiktok")) copySheet.appendRow(["copy_link_tiktok", "TikTok Link", "https://www.tiktok.com/@up2.media"]);
     if(!ids.includes("copy_link_fb")) copySheet.appendRow(["copy_link_fb", "Facebook Link", "https://www.facebook.com"]);
     if(!ids.includes("copy_link_threads")) copySheet.appendRow(["copy_link_threads", "Threads Link", "https://www.threads.com/@up2.media"]);
     if(!ids.includes("copy_link_blogger")) copySheet.appendRow(["copy_link_blogger", "Blogger Link", "https://kulineriadigital.blogspot.com/"]);
  }
  
  // Tambahkan Sheet "Settings" otomatis untuk menyimpan PIN
  let settingSheet = ss.getSheetByName("Settings");
  if (!settingSheet) {
     settingSheet = ss.insertSheet("Settings");
     settingSheet.appendRow(["Key_Configuration", "Value_Data"]);
     settingSheet.appendRow(["ADMIN_PIN", "1234"]);
  }
  
  Logger.log("==== PATCH DATABASE SELESAI ====");
}