const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];

const STORAGE_KEYS = { cart: "luzita_cart", favorites: "luzita_favorites" };
const COLOR_MAP = {
  أبيض: "#fff",
  "بيج فاتح": "#eadfcb",
  "وردي فاتح": "#f4d3d6",
  سماوي: "#cfe5f2",
  لافندر: "#ddd3f0",
  نعناعي: "#d3ecd9",
  ليموني: "#f6efc4",
  خوخي: "#f9dcc8",
  أسود: "#171514",
  رمادي: "#aaa29d",
  كحلي: "#17243a",
  زيتي: "#4f5b3b",
  بني: "#74523f",
  أزرق: "#507aa3",
  أخضر: "#4d7456",
  وردي: "#d98d9b",
  نبيتي: "#6e2730",
};
const state = {
  products: [],
  categories: [],
  settings: {},
  cart: JSON.parse(localStorage.getItem(STORAGE_KEYS.cart) || "[]"),
  favorites: JSON.parse(localStorage.getItem(STORAGE_KEYS.favorites) || "[]"),
};

const money = (value) =>
  Math.round(Number(value) || 0).toLocaleString("ar-EG") + " جنيه";
const save = () => {
  localStorage.setItem(STORAGE_KEYS.cart, JSON.stringify(state.cart));
  localStorage.setItem(STORAGE_KEYS.favorites, JSON.stringify(state.favorites));
  updateCounters();
};
const toast = (message) => {
  const el = $("#toast");
  el.textContent = message;
  el.classList.add("show");
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => el.classList.remove("show"), 2200);
};
const imageFallback = (img) => {
  img.onerror = () => {
    img.onerror = null;
    img.src =
      "https://images.unsplash.com/photo-1529139574466-a303027c1d8b?w=900";
  };
};

async function loadExcel() {
  try {
    const response = await fetch("products.xlsx");
    if (!response.ok) throw new Error("Excel not found");
    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const productRows = XLSX.utils.sheet_to_json(workbook.Sheets.Products, {
      defval: "",
    });
    const categoryRows = XLSX.utils.sheet_to_json(workbook.Sheets.Categories, {
      defval: "",
    });
    const settingRows = XLSX.utils.sheet_to_json(workbook.Sheets.Settings, {
      defval: "",
    });
    state.products = productRows.map((p, i) => ({
      ...p,
      id: i + 1,
      pack_qty: +p.pack_qty,
      pack_price: +p.pack_price,
      stock_pieces: +p.stock_pieces,
      rating: +p.rating,
      reviews: +p.reviews,
      colors: String(p.colors)
        .split(",")
        .map((x) => x.trim()),
      sizes: String(p.sizes)
        .split(",")
        .map((x) => x.trim()),
    }));
    state.categories = categoryRows;
    state.settings = Object.fromEntries(
      settingRows.map((x) => [x.key, x.value]),
    );
  } catch (error) {
    console.error(error);
    toast("تعذر قراءة products.xlsx");
  }
  updateCounters();
  fillFooter();
  route();
}

function updateCounters() {
  $("#cartCount").textContent = state.cart.reduce(
    (sum, item) => sum + item.quantity,
    0,
  );
  $("#favCount").textContent = state.favorites.length;
}
function fillFooter() {
  $("#footerCats").innerHTML = state.categories
    .map(
      (c) =>
        `<a href="#/products?cat=${encodeURIComponent(c.category)}">${c.title}</a>`,
    )
    .join("");
  const s = state.settings;
  $("#contactInfo").innerHTML =
    `واتساب: ${s.whatsapp || "—"}<br>الهاتف: ${s.phone || "—"}<br>البريد: ${s.email || "—"}`;
}
function productCard(product) {
  const favorite = state.favorites.includes(product.id);
  const image = product.image_url;
  console.log(product.image_url);
  return `<article class="product">
    <div class="product-img"><img src="${image}" alt="${product.name}" loading="lazy" onerror="this.src='https://images.unsplash.com/photo-1529139574466-a303027c1d8b?w=900'">
      <span class="badge">${product.collection}</span>
      <button class="heart" data-fav="${product.id}" aria-label="المفضلة">${favorite ? "♥" : "♡"}</button>
    </div>
    <div class="product-body">
      <h3>${product.name}</h3><div class="meta">${product.category} • ${product.audience} • ${product.code}</div>
      <div class="price-row"><div><div class="price">${money(product.pack_price)}</div><div class="piece">${product.pack_qty} قطعة في العبوة • ${money(product.pack_price / product.pack_qty)} للقطعة</div></div></div>
      <div class="rating">★ ${product.rating} <span class="meta">(${product.reviews} تقييم)</span></div>
      <div class="colors">${product.colors.map((c) => `<i title="${c}" style="background:${COLOR_MAP[c] || c}"></i>`).join("")}</div>
      <div class="product-actions"><a class="btn outline" href="#/product/${product.id}">التفاصيل</a></div>
    </div>
  </article>`;
}
function home() {
  const categories = state.categories
    .map(
      (c) =>
        `<a class="category" href="#/products?cat=${encodeURIComponent(c.category)}"><img src="${c.image_url}" alt="${c.title}" loading="lazy"><div><h3>${c.title}</h3><p>${c.description}</p></div></a>`,
    )
    .join("");
  const products = state.products.slice(0, 8).map(productCard).join("");
  const top = [...state.products]
    .sort((a, b) => b.reviews - a.reviews)
    .slice(0, 4)
    .map(productCard)
    .join("");
  $("#app").innerHTML =
    `<section class="hero"><div class="container hero-inner">
    <div class="hero-copy"><div class="eyebrow">للتجار وأصحاب المحلات</div><h1>ملابس بالجملة<br><span>بعبوات واضحة</span></h1>
    <p>تشكيلات رجالي وحريمي وشبابي وأطفال، مع سعر العبوة وسعر القطعة واضحين قبل الطلب.</p>
    <div class="actions"><a class="btn" href="#/products">تسوق بالجملة</a><a class="btn outline" href="#/products">استكشف المنتجات</a></div></div>
    <div class="hero-photo"><img src="images/hero/photo-1483985988355-763728e1935b.jfif" alt="تشكيلة ملابس لوزيتا"></div>
  </div></section>
  <section class="section"><div class="container"><div class="section-head"><h2>التصنيفات الرئيسية</h2><a href="#/products">عرض الكل ←</a></div><div class="categories">${categories}</div></div></section>
  <section class="section alt"><div class="container"><div class="section-head"><h2>أفضل المنتجات</h2><a href="#/products">عرض الكل ←</a></div><div class="products">${products}</div></div></section>
  <section class="section"><div class="container"><div class="features"><div class="feature-photo"></div><div class="feature-list">
    <div class="feature"><b>أسعار عبوات واضحة</b><p>اعرف سعر العبوة وسعر القطعة مباشرة.</p></div>
    <div class="feature"><b>تنوع كبير</b><p>ألوان ومقاسات وتشكيلات متعددة.</p></div>
    <div class="feature"><b>طلب جملة سهل</b><p>حدد عدد العبوات وشاهد الإجمالي فورًا.</p></div>
    <div class="feature"><b>خدمة عملاء</b><p>تواصل معنا عبر واتساب بسهولة.</p></div>
  </div></div></div></section>
  <section class="section alt"><div class="container"><div class="section-head"><h2>الأكثر مبيعًا</h2><a href="#/products?sort=popular">عرض الكل ←</a></div><div class="products">${top}</div></div></section>
  <section class="section"><div class="container"><div class="section-head"><h2>آراء العملاء</h2></div><div class="testimonials">
    <article class="quote"><div class="stars">★★★★★</div><b>محل الأناقة</b><p>العبوات واضحة والحساب سهل جدًا، وده وفر علينا وقت كبير.</p></article>
    <article class="quote"><div class="stars">★★★★★</div><b>تاجر جملة</b><p>تنوع جيد في الألوان والمقاسات والأسعار مناسبة للتجار.</p></article>
    <article class="quote"><div class="stars">★★★★☆</div><b>بوتيك الريم</b><p>اختيار المنتجات والطلب بالجملة بسيط وواضح.</p></article>
  </div></div></section>
  <section class="section"><div class="container contact-box"><div><div class="eyebrow">نحن جاهزون للرد</div><h2>عايز تشكيلة أو سعر خاص؟</h2><p>ابعتلنا احتياجك على واتساب ونساعدك تختار الكميات المناسبة.</p><a class="btn" id="whatsappBtn">راسلنا على واتساب</a></div>
  <form class="form" id="contactForm"><input id="contactName" placeholder="الاسم" required><textarea id="contactMessage" rows="4" placeholder="اكتب رسالتك" required></textarea><button class="btn">إرسال عبر واتساب</button></form></div></section>`;
}
function productsPage() {
  const params = new URLSearchParams(location.hash.split("?")[1] || "");
  const category = params.get("cat") || "";
  const sort = params.get("sort") || "";
  let list = state.products.filter(
    (p) => !category || p.audience === category || p.category === category,
  );
  if (sort === "popular") list.sort((a, b) => b.reviews - a.reviews);
  $("#app").innerHTML =
    `<section class="page"><div class="container"><div class="section-head"><div><div class="eyebrow">لوزيتا</div><h1>كل المنتجات</h1></div></div>
  <div class="filter"><input id="productSearch" placeholder="ابحث باسم المنتج أو الرمز"><select id="categoryFilter"><option value="">كل التصنيفات</option>${state.categories.map((c) => `<option>${c.category}</option>`).join("")}</select><select id="collectionFilter"><option value="">كل المجموعات</option><option>اقتصادي وشعبي</option><option>تجاري يومي</option><option>عالي الجودة</option><option>مميز وفاخر</option></select></div>
  <div class="products" id="productList">${list.map(productCard).join("")}</div></div></section>`;
  $("#categoryFilter").value = category;
  const render = () => {
    const q = $("#productSearch").value.trim().toLowerCase(),
      cat = $("#categoryFilter").value,
      col = $("#collectionFilter").value;
    const r = state.products.filter(
      (p) =>
        (!q || (p.name + p.code + p.category).toLowerCase().includes(q)) &&
        (!cat || p.audience === cat || p.category === cat) &&
        (!col || p.collection === col),
    );
    $("#productList").innerHTML = r.length
      ? r.map(productCard).join("")
      : `<div class="empty">لا توجد منتجات مطابقة.</div>`;
  };
  $$(".filter input,.filter select").forEach((el) =>
    el.addEventListener("input", render),
  );
}

function detailPage(id) {
  const p = state.products.find((x) => x.id == id);

  if (!p) {
    $("#app").innerHTML = `
      <section class="page">
        <div class="container empty">المنتج غير موجود.</div>
      </section>`;
    return;
  }

  const sizes = p.sizes.filter(Boolean);

  const safeSizes = sizes.length ? sizes : ["مقاس موحد"];

  let selectedSize = safeSizes[0];

  $("#app").innerHTML = `
    <section class="page">
      <div class="container detail">

        <div class="detail-image">
          <img src="${p.image_url}" alt="${p.name}"
               onerror="this.onerror=null;this.src='https://images.unsplash.com/photo-1529139574466-a303027c1d8b?w=900'">
        </div>

        <div>
          <div class="eyebrow">${p.collection}</div>
          <h1>${p.name}</h1>

          <div class="meta">
            ${p.code} • ${p.category} • ${p.audience}
          </div>

          <div class="detail-price">${money(p.pack_price)}</div>

          <p>
            سعر العبوة — ${p.pack_qty} قطعة.
            <br>
            سعر القطعة: ${money(p.pack_price / p.pack_qty)}
          </p>

          <div class="option-group">
            <div class="option-title">
              <span>اختر المقاس</span>
              <small id="selectedSizeLabel">${selectedSize}</small>
            </div>

            <div class="size-options" id="sizeOptions">
              ${safeSizes
                .map(
                  (size, i) => `
                <button type="button"
                  class="size-option ${i === 0 ? "selected" : ""}"
                  data-size-index="${i}">
                  ${size}
                </button>
              `,
                )
                .join("")}
            </div>
          </div>

          

            

          <div class="quantity">
            <button type="button" id="minus" aria-label="تقليل العبوات">−</button>
            <input id="quantity" type="number" min="1" step="1" value="1">
            <button type="button" id="plus" aria-label="زيادة العبوات">+</button>
          </div>

          <div class="cart-options" id="quantitySummary">
            عدد القطع: ${p.pack_qty} قطعة
          </div>

          <div class="cart-options" id="priceSummary">
            الإجمالي: ${money(p.pack_price)}
          </div>

          <div id="selectionError" class="selection-error" role="alert"></div>

          <button type="button" class="btn" id="addDetail">
            أضف للسلة
          </button>
        </div>

      </div>
    </section>`;

  const sizeButtons = $$("#sizeOptions [data-size-index]");
  const colorButtons = $$("#colorOptions [data-color-index]");

  sizeButtons.forEach((button) => {
    button.addEventListener("click", () => {
      sizeButtons.forEach((b) => b.classList.remove("selected"));
      button.classList.add("selected");

      selectedSize = safeSizes[Number(button.dataset.sizeIndex)];
      $("#selectedSizeLabel").textContent = selectedSize;
      $("#selectionError").textContent = "";
    });
  });

  colorButtons.forEach((button) => {
    button.addEventListener("click", () => {
      colorButtons.forEach((b) => b.classList.remove("selected"));
      button.classList.add("selected");

      selectedColor = safeColors[Number(button.dataset.colorIndex)];
      $("#selectedColorLabel").textContent = selectedColor;
      $("#selectionError").textContent = "";
    });
  });

  const quantityInput = $("#quantity");

  function updateSummary() {
    let quantity = Math.floor(Number(quantityInput.value));

    if (!Number.isFinite(quantity) || quantity < 1) {
      quantity = 1;
      quantityInput.value = 1;
    }

    $("#quantitySummary").textContent =
      `عدد القطع: ${quantity * p.pack_qty} قطعة`;

    $("#priceSummary").textContent =
      `الإجمالي: ${money(p.pack_price * quantity)}`;
  }

  $("#minus").onclick = () => {
    quantityInput.value = Math.max(1, (Number(quantityInput.value) || 1) - 1);
    updateSummary();
  };

  $("#plus").onclick = () => {
    quantityInput.value = Math.max(
      1,
      Math.floor(Number(quantityInput.value) || 1) + 1,
    );
    updateSummary();
  };

  quantityInput.addEventListener("input", updateSummary);

  $("#addDetail").onclick = () => {
    const quantity = Math.floor(Number(quantityInput.value));

    if (!Number.isFinite(quantity) || quantity < 1) {
      $("#selectionError").textContent = "من فضلك أدخل عدد عبوات صحيح.";
      return;
    }

    addToCart(id, quantity, selectedSize);
  };
}

function addToCart(id, quantity = 1, size = "") {
  const p = state.products.find((x) => x.id == id);
  if (!p) return;

  quantity = Math.floor(Number(quantity));

  if (!Number.isFinite(quantity) || quantity < 1) {
    toast("من فضلك حدد عدد عبوات صحيح");
    return;
  }

  const item = state.cart.find((x) => x.productId === p.id && x.size === size);

  if (item) {
    item.quantity += quantity;
  } else {
    state.cart.push({
      productId: p.id,
      quantity,
      size,
    });
  }

  save();
  toast("تمت إضافة المنتج للسلة");
}

function toggleFavorite(id) {
  state.favorites = state.favorites.includes(id)
    ? state.favorites.filter((x) => x !== id)
    : [...state.favorites, id];
  save();
  route();
}
function cartPage() {
  const rows = state.cart
    .map((item) => {
      const p = state.products.find((x) => x.id === item.productId);
      return p ? { ...item, p } : null;
    })
    .filter(Boolean);
  const total = rows.reduce((s, x) => s + x.p.pack_price * x.quantity, 0),
    pieces = rows.reduce((s, x) => s + x.p.pack_qty * x.quantity, 0);
  $("#app").innerHTML =
    `<section class="page"><div class="container"><div class="section-head"><h1>سلة التسوق</h1></div>${
      rows.length
        ? `<div class="products">${rows
            .map(
              (
                x,
              ) => `<article class="product"><div class="product-img"><img src="${x.p.image_url}" alt="${x.p.name}"></div><div class="product-body"><h3>${x.p.name}</h3> <div class="meta">${x.p.pack_qty} قطعة في العبوة</div>
<div class="cart-options">
  المقاس: ${x.size || "غير محدد"}
</div> <div class="price">${money(x.p.pack_price * x.quantity)}</div><div class="quantity"><button data-dec="${x.productId}">−</button><b>${x.quantity}</b><button data-inc="${x.productId}">+</button></div><button class="btn outline" data-remove="${x.productId}|${x.size || ""}">حذف</button></div></article>`,
            )
            .join(
              "",
            )}</div><div class="contact-box" style="margin-top:25px"><div><h2>ملخص الطلب</h2><p>إجمالي القطع: ${pieces}</p><p>الإجمالي: <strong>${money(total)}</strong></p></div><button class="btn" id="checkout">إرسال الطلب عبر واتساب</button></div>`
        : `<div class="empty">السلة فارغة. <a href="#/products" style="color:var(--primary)">تصفح المنتجات</a></div>`
    }</div></section>`;
}
function sendWhatsApp(message) {
  const number = state.settings.whatsapp || "";
  if (!number) {
    toast("أضف رقم واتساب في شيت Settings");
    return;
  }
  window.open(
    `https://wa.me/${number}?text=${encodeURIComponent(message)}`,
    "_blank",
  );
}

function checkout() {
  const rows = state.cart
    .map((x) => ({
      ...x,
      p: state.products.find((y) => y.id === x.productId),
    }))
    .filter((x) => x.p);

  if (!rows.length) return;

  const total = rows.reduce((sum, x) => sum + x.p.pack_price * x.quantity, 0);

  const deposit = total * 0.1;
  const depositNumber = state.settings.deposit_number || "";

  document.getElementById("checkoutModal")?.remove();

  const modal = document.createElement("div");
  modal.id = "checkoutModal";
  modal.className = "checkout-overlay";

  modal.innerHTML = `
    <div class="checkout-modal" dir="rtl">
      <button type="button" class="checkout-close" id="closeCheckout"
        aria-label="إغلاق">×</button>

      <h2>تأكيد بيانات الطلب</h2>
      <p class="checkout-description">
        أدخل بياناتك وعنوان التوصيل لمتابعة طلبك.
      </p>

      <div class="deposit-summary">
        <div class="deposit-row">
          <span>إجمالي الفاتورة</span>
          <strong>${money(total)}</strong>
        </div>

        <div class="deposit-row">
          <span>العربون المطلوب (10%)</span>
          <strong>${money(deposit)}</strong>
        </div>

        <div class="deposit-number">
          <span>رقم تحويل العربون</span>
          <strong id="depositNumber"></strong>
        </div>

        <p>
          لتأكيد الطلب، حوّل العربون الموضح أعلاه إلى الرقم المحدد.
          بعد تأكيد الطلب من الموقع، أرسل صورة التحويل على واتساب المتجر
          مع تفاصيل طلبك.
        </p>
      </div>

      <form id="checkoutForm" class="checkout-form">
        <label>
          الاسم بالكامل *
          <input name="customerName" required maxlength="100"
            placeholder="اكتب اسمك بالكامل">
        </label>

        <label>
          رقم الهاتف *
          <input name="customerPhone" required type="tel"
            maxlength="20" placeholder="رقم الهاتف للتواصل">
        </label>

        <label>
          المحافظة / المدينة *
          <input name="city" required maxlength="100"
            placeholder="مثال: القاهرة">
        </label>

        <label>
          عنوان التوصيل بالتفصيل *
          <textarea name="address" required rows="3" maxlength="500"
            placeholder="الشارع، رقم العقار، الدور، وعلامة مميزة"></textarea>
        </label>

        <label>
          ملاحظات إضافية
          <textarea name="notes" rows="2" maxlength="500"
            placeholder="أي تفاصيل إضافية (اختياري)"></textarea>
        </label>

        <p class="checkout-warning">
          تنبيه: قيمة العربون 10% من إجمالي الفاتورة. تأكد من التحويل
          للرقم الموضح، واحتفظ بصورة التحويل لإرسالها على واتساب المتجر.
        </p>

        <div class="checkout-actions">
          <button type="submit" class="btn">
            تأكيد البيانات والمتابعة إلى واتساب
          </button>
          <button type="button" class="btn btn-secondary"
            id="cancelCheckout">رجوع للسلة</button>
        </div>
      </form>
    </div>
  `;

  document.body.appendChild(modal);

  document.getElementById("depositNumber").textContent =
    depositNumber || "لم يتم إعداد رقم التحويل";

  document.getElementById("closeCheckout").onclick = document.getElementById(
    "cancelCheckout",
  ).onclick = () => modal.remove();

  modal.addEventListener("click", (e) => {
    if (e.target === modal) modal.remove();
  });
}

function route() {
  const hash = location.hash.slice(1) || "/";
  const [path] = hash.split("?");
  if (path === "/") home();
  else if (path === "/products") productsPage();
  else if (path.startsWith("/product/")) detailPage(path.split("/")[2]);
  else if (path === "/cart") cartPage();
  else if (path === "/about")
    $("#app").innerHTML =
      `<section class="page"><div class="container empty"><h1>من نحن</h1><p>لوزيتا متجر متخصص في بيع الملابس بالجملة للتجار وأصحاب المحلات.</p></div></section>`;
  else if (path === "/contact") home();
  else if (path === "/favs") {
    const list = state.products.filter((p) => state.favorites.includes(p.id));
    $("#app").innerHTML =
      `<section class="page"><div class="container"><h1>المفضلة</h1><div class="products">${list.map(productCard).join("") || "<div class='empty'>لا توجد منتجات في المفضلة.</div>"}</div></div></section>`;
  } else home();
  updateCounters();
  window.scrollTo({ top: 0, behavior: "smooth" });
}
document.addEventListener("click", (e) => {
  const fav = e.target.closest("[data-fav]"),
    add = e.target.closest("[data-add]"),
    inc = e.target.closest("[data-inc]"),
    dec = e.target.closest("[data-dec]"),
    remove = e.target.closest("[data-remove]");
  if (fav) toggleFavorite(+fav.dataset.fav);
  if (add) addToCart(+add.dataset.add);
  if (inc) {
    const x = state.cart.find((x) => x.productId == inc.dataset.inc);
    if (x) x.quantity++;
    save();
    cartPage();
  }
  if (dec) {
    const x = state.cart.find((x) => x.productId == dec.dataset.dec);
    if (x) {
      x.quantity = Math.max(1, x.quantity - 1);
      save();
      cartPage();
    }
  }
  if (remove) {
    const [productId, size] = remove.dataset.remove.split("|");

    state.cart = state.cart.filter(
      (x) => !(String(x.productId) === productId && x.size === size),
    );

    save();
    cartPage();
  }
  if (e.target.id === "checkout") checkout();
});
$("#searchForm").onsubmit = (e) => {
  e.preventDefault();
  location.hash = "#/products?q=" + encodeURIComponent($("#searchInput").value);
};

document.addEventListener("submit", function (e) {
  if (e.target.id !== "checkoutForm") return;

  e.preventDefault();

  const form = e.target;
  if (!form.reportValidity()) return;

  const whatsapp = state.settings.whatsapp || "";
  const depositNumber = state.settings.deposit_number || "";

  if (!whatsapp) {
    toast("أضف رقم واتساب المتجر في شيت Settings");
    return;
  }

  if (!depositNumber) {
    toast("أضف رقم تحويل العربون في شيت Settings");
    return;
  }

  const data = new FormData(form);

  const customerName = data.get("customerName").trim();
  const customerPhone = data.get("customerPhone").trim();
  const city = data.get("city").trim();
  const address = data.get("address").trim();
  const notes = data.get("notes").trim();

  const rows = state.cart
    .map((x) => ({
      ...x,
      p: state.products.find((y) => y.id === x.productId),
    }))
    .filter((x) => x.p);

  if (!rows.length) {
    toast("السلة فارغة");
    return;
  }

  const total = rows.reduce((sum, x) => sum + x.p.pack_price * x.quantity, 0);

  const deposit = total * 0.1;

  const message = [
    "طلب جديد من لوزيتا",
    "",
    "بيانات العميل:",
    `الاسم: ${customerName}`,
    `رقم الهاتف: ${customerPhone}`,
    `المحافظة / المدينة: ${city}`,
    `عنوان التوصيل: ${address}`,
    notes ? `ملاحظات: ${notes}` : "",
    "",
    "تفاصيل الطلب:",
    ...rows.map(
      (x) =>
        `- ${x.p.name}
المقاس: ${x.size || "غير محدد"}
الكمية: ${x.quantity} عبوة × ${x.p.pack_qty} قطعة
السعر: ${money(x.p.pack_price * x.quantity)}`,
    ),
    "",
    `إجمالي الفاتورة: ${money(total)}`,
    `العربون المطلوب (10%): ${money(deposit)}`,
    `رقم تحويل العربون: ${depositNumber}`,
    "",
    "تعليمات العربون:",
    "يرجى تحويل العربون إلى الرقم الموضح أعلاه.",
    "بعد تأكيد الطلب من الموقع، سأرسل صورة التحويل على واتساب المتجر.",
  ]
    .filter(Boolean)
    .join("\n");

  const modal = document.getElementById("checkoutModal");

  window.open(
    `https://wa.me/${whatsapp}?text=${encodeURIComponent(message)}`,
    "_blank",
  );

  modal?.remove();
});

$("#menuBtn").onclick = () => $("#nav").classList.toggle("open");
window.addEventListener("hashchange", route);
loadExcel();
