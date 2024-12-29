const puppeteer = require("puppeteer");
const xlsx = require("xlsx");
require("dotenv").config();

// Hàm chờ theo số mili giây
function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

(async () => {
  // 1. Đọc file Excel orderLV.xlsx
  console.log("[INFO] Reading Excel file: orderLV.xlsx");
  const workbook = xlsx.readFile("orderLV.xlsx");
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const orders = xlsx.utils.sheet_to_json(sheet);

  console.log("[INFO] Orders loaded:", orders);

    // 2. Mở trình duyệt
    console.log("[INFO] Launching browser with full screen...");
    const browser = await puppeteer.launch({
      headless: false,
      defaultViewport: null,
      args: ["--start-maximized"],
    });
    const page = await browser.newPage();
  
    // Đặt User-Agent giống Chrome thật
    const customUserAgent =
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
      "AppleWebKit/537.36 (KHTML, like Gecko) " +
      "Chrome/131.0.0.0 Safari/537.36";
    console.log(`[INFO] Setting user agent: ${customUserAgent}`);
    await page.setUserAgent(customUserAgent);
  
    // 3. Truy cập trang chủ LV
    const homeUrl = "https://uk.louisvuitton.com/eng-gb/homepage";
    console.log(`[INFO] Going to homepage: ${homeUrl}`);
    try {
      await page.goto(homeUrl, {
        waitUntil: "domcontentloaded",
        timeout: 60000,
      });
      console.log("[INFO] Homepage loaded successfully.");
    } catch (err) {
      console.error("[ERROR] Failed to load LV homepage:", err.message);
    }
  
    // 4. Set token đăng nhập (nếu cần)
    const loginKey = "LV.loginData";
    const loginValue = JSON.stringify({
      usid: "a48001fe-9796-4e7b-98ec-afd3a26ea788",
    });
    console.log("[INFO] Setting token in localStorage...");
    await page.evaluate((key, value) => localStorage.setItem(key, value), loginKey, loginValue);
  
    console.log("[INFO] Reloading page to apply token...");
    try {
      await page.reload({ waitUntil: "domcontentloaded", timeout: 60000 });
      console.log("[INFO] Token LV.loginData set successfully.");
    } catch (err) {
      console.error("[ERROR] Reload page failed after setting token:", err.message);
    }

  // **Selector** tùy theo info bạn cung cấp
  const searchButtonSelector = "button#headerSearchButton"; // Nút search
  const searchInputSelector = ".lv-search-input__input";
  const productItemSelector = ".lv-smart-link.lv-product-card__url";
  const addToCartSelector = ".lv-product-purchase-button";
  const cartIconSelector = ".lv-header-icon-shopping-bag";
  const cartLinkSelector = ".lv-smart-link";

  // 5. Lặp qua từng đơn hàng
  for (const order of orders) {
    const { SKU, ProductID, Name, Quantity = 1 } = order;
    console.log(`\n--- [ORDER] SKU: ${SKU}, Name: ${Name}, ProductID: ${ProductID}, Qty: ${Quantity} ---`);

    // Tắt thông báo Cookie nếu xuất hiện
    const cookieCloseSelector = ".ucm-popin-close-text";

    console.log("[INFO] Checking for cookie preferences popup...");
    try {
      await page.waitForSelector(cookieCloseSelector, { timeout: 5000 });
      console.log("[INFO] Cookie preferences popup found, closing...");
      await page.click(cookieCloseSelector);
      console.log("[INFO] Cookie preferences popup closed successfully.");
    } catch (err) {
      console.warn("[WARN] Cookie preferences popup not found. Skipping...");
    }

        // Tắt popup Select Country nếu xuất hiện
    const countryModalCloseSelector = "button.lv-modal__close";

    console.log("[INFO] Checking for country selection modal...");
    try {
      await page.waitForSelector(countryModalCloseSelector, { timeout: 5000 });
      console.log("[INFO] Country selection modal found, closing...");
      await page.click(countryModalCloseSelector);
      console.log("[INFO] Country selection modal closed successfully.");
    } catch (err) {
      console.warn("[WARN] Country selection modal not found. Skipping...");
    }


    console.log(`[STEP] A: Waiting for search button (${searchButtonSelector})...`);
    try {
      // Chờ nút search xuất hiện
      await page.waitForSelector(searchButtonSelector, { timeout: 10000 });
      console.log("[INFO] Search button found, clicking...");

      // Bấm nút search
      await page.click(searchButtonSelector);

      // Debug nếu nhấn không thành công
      console.log("[INFO] Search button clicked successfully.");
    } catch (err) {
      console.error(`[ERROR] Search button (${searchButtonSelector}) not found or clickable.`, err.message);

      // Thử nhấn nút bằng JavaScript (dự phòng nếu click() thất bại)
      try {
        console.log("[INFO] Attempting to click search button via JavaScript...");
        await page.evaluate((selector) => {
          document.querySelector(selector)?.click();
        }, searchButtonSelector);
        console.log("[INFO] Search button clicked via JavaScript successfully.");
      } catch (jsErr) {
        console.error(`[ERROR] Failed to click search button via JavaScript.`, jsErr.message);
        continue; // Bỏ qua đơn hàng nếu nút Search không thể nhấn
      }
    }

    // B. Nhập tên sản phẩm vào ô search
    console.log(`[STEP] B: Waiting for search input (${searchInputSelector}) after clicking search button...`);
    try {
      await page.waitForSelector(searchInputSelector, { timeout: 10000 });
      console.log("[INFO] Search input now visible, typing product Name...");
      // Xóa nội dung cũ (nếu có)
      await page.click(searchInputSelector, { clickCount: 3 });
      await page.type(searchInputSelector, Name, { delay: 100 });
      // Nhấn Enter
      console.log("[INFO] Pressing Enter to search...");
      await page.keyboard.press("Enter");
    } catch (err) {
      console.error(`[ERROR] Search input (${searchInputSelector}) not found after clicking search button.`, err.message);
      continue; // skip this order
    }

    // C. Đợi kết quả load
    console.log("[STEP] C: Waiting 3 seconds for search results...");
    await delay(3000);

    // D. Tìm link sản phẩm theo ProductID
    console.log(`[STEP] D: Checking product items with selector (${productItemSelector})...`);
    try {
      await page.waitForSelector(productItemSelector, { timeout: 15000 });
      console.log("[INFO] Product list found, retrieving hrefs...");

      // Lấy danh sách link
      const productLinks = await page.$$eval(productItemSelector, (els) =>
        els.map((el) => el.getAttribute("href"))
      );

      console.log("[INFO] Found product links:\n", productLinks);

      // Tìm link có chứa ProductID
      const foundLink = productLinks.find((link) => link.includes(ProductID));
      if (!foundLink) {
        console.warn(`[WARN] No product link found with ProductID: ${ProductID}`);
        continue;
      }
      console.log(`[INFO] Found link with ProductID: ${foundLink}`);

      // Xác định index
      const index = productLinks.indexOf(foundLink);
      // Tìm element tương ứng để click
      const productElements = await page.$$(productItemSelector);
      const targetElement = productElements[index];
      console.log(`[INFO] Clicking product item at index: ${index}`);
      await targetElement.click();
    } catch (err) {
      console.error(`[ERROR] Failed to find/click product link for ProductID: ${ProductID}`, err.message);
      continue;
    }

    // E. Thêm vào giỏ
    console.log(`[STEP] E: Delaying 5s for product detail to load...`);
    await delay(5000);

    console.log(`[STEP] E: Trying to click Add To Cart button (${addToCartSelector})...`);
    try {
      await page.waitForSelector(addToCartSelector, { timeout: 10000 });
      await page.click(addToCartSelector);
      console.log(`[INFO] SKU: ${SKU} - Added to cart successfully.`);
    } catch (err) {
      console.error(`[ERROR] Add To Cart button (${addToCartSelector}) not found for SKU: ${SKU}.`, err.message);
      continue;
    }

    // // F. Mở giỏ hàng
    // console.log("[STEP] F: Delaying 3s then opening cart...");
    // await delay(3000);

    // console.log(`[STEP] F: Waiting for cart icon (${cartIconSelector})...`);
    // try {
    //   await page.waitForSelector(cartIconSelector, { timeout: 10000 });
    //   console.log("[INFO] Cart icon found, clicking...");
    //   await page.click(cartIconSelector);

    //   console.log(`[INFO] Waiting for cart link (${cartLinkSelector})...`);
    //   await page.waitForSelector(cartLinkSelector, { timeout: 10000 });
    //   console.log("[INFO] Cart link found, clicking...");
    //   await page.click(cartLinkSelector);
    //   console.log(`[INFO] SKU: ${SKU} - Cart page opened successfully.`);
    // } catch (err) {
    //   console.error(`[ERROR] Cannot open cart for SKU: ${SKU}.`, err.message);
    //   continue;
    // }

    // // G. Điều chỉnh số lượng (nếu có)
    // console.log("[STEP] G: Delaying 3s before adjusting quantity...");
    // await delay(3000);
    // for (let i = 1; i < Quantity; i++) {
    //   // Tuỳ website có hỗ trợ + / - không
    //   // Vd: await page.click(".quantity-plus-btn");
    //   // await delay(1000);
    // }
    // console.log(`[INFO] SKU: ${SKU} - Quantity set to ${Quantity} (if site supports).`);

    // console.log(`[SUCCESS] Completed order for SKU: ${SKU}.`);
  }

  // 6. Đóng trình duyệt
  // console.log("[INFO] Closing browser...");
  // await browser.close();
  console.log("\n=== [DONE] All orders processed ===");
})();
