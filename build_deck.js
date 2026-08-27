// Bài trình chiếu: Lịch sử tỉnh Quảng Ninh
// Dựng bằng pptxgenjs — layout 16:9 rộng (13.333 x 7.5 inch)
const pptxgen = require("pptxgenjs");
const p = new pptxgen();
p.layout = "LAYOUT_WIDE";
p.author = "Arena Agent";
p.company = "XroHub";
p.title = "Lịch sử tỉnh Quảng Ninh";
p.subject = "Lịch sử dựng nước và giữ nước gắn với tỉnh Quảng Ninh";

const W = 13.333, H = 7.5;
const C = {
  dark:   "0A2E26",   // xanh lục bảo đậm (chủ đạo)
  darker: "061F19",
  deep:   "0E3B30",
  jade:   "12715A",   // ngọc bích
  jadeLt: "2E9E80",
  gold:   "E3A83B",   // vàng đồng (nhấn)
  goldLt: "F2CE84",
  cream:  "F7F2E7",   // kem (nền sáng)
  cream2: "EFE7D6",
  ink:    "20302A",
  gray:   "5E6E66",
  white:  "FFFFFF",
};
const F = "Arial";
const IMG = "images/";

function sh(){ return { type: "outer", color: "123028", opacity: 0.22, blur: 9, offset: 3, angle: 90 }; }

function fullBg(slide, color){
  slide.addShape(p.shapes.RECTANGLE, { x:0, y:0, w:W, h:H, fill:{ color } });
}

// Triện chương: vòng tròn vàng đồng mang số thứ tự
function seal(slide, x, y, d, num, opts = {}){
  slide.addShape(p.shapes.OVAL, {
    x, y, w:d, h:d,
    fill: { color: opts.fill || C.gold },
    line: opts.ring ? { color: opts.ring, width: 1.5 } : undefined,
  });
  slide.addText(num, {
    x, y: y + d*0.06, w:d, h:d*0.88,
    align: "center", valign: "middle", margin: 0,
    fontFace: F, bold: true, fontSize: opts.size || d*30,
    color: opts.textColor || C.dark,
  });
}

// Header cho slide chương
function chapterHeader(slide, num, kicker, title, era){
  slide.addText(kicker, {
    x:0.55, y:0.40, w:10.2, h:0.32, margin:0,
    fontFace:F, fontSize:12.5, bold:true, color:C.jade, charSpacing:2,
  });
  slide.addText(title, {
    x:0.55, y:0.72, w:10.7, h:0.72, margin:0,
    fontFace:F, fontSize:29, bold:true, color:C.dark,
  });
  slide.addShape(p.shapes.RECTANGLE, { x:0.57, y:1.47, w:1.05, h:0.05, fill:{ color:C.gold } });
  const eraW = 0.3 + era.length * 0.082;
  slide.addShape(p.shapes.ROUNDED_RECTANGLE, {
    x:1.82, y:1.335, w:eraW, h:0.34, rectRadius:0.17,
    fill:{ color:C.dark },
  });
  slide.addText(era, {
    x:1.82, y:1.335, w:eraW, h:0.34, margin:0,
    align:"center", valign:"middle",
    fontFace:F, fontSize:11, bold:true, color:C.goldLt, charSpacing:1,
  });
  seal(slide, 12.25, 0.42, 0.88, num, { size:26 });
}

function footer(slide, page){
  slide.addText("LỊCH SỬ TỈNH QUẢNG NINH", {
    x:0.55, y:7.14, w:5, h:0.26, margin:0,
    fontFace:F, fontSize:8.5, color:C.gray, charSpacing:2,
  });
  slide.addText(String(page), {
    x:12.4, y:7.14, w:0.4, h:0.26, margin:0, align:"right",
    fontFace:F, fontSize:9, bold:true, color:C.jade,
  });
}

// Ảnh bo góc + chú thích
function photo(slide, file, x, y, w, h, caption){
  slide.addImage({ path: IMG + file, x, y, w, h, rounding: true });
  slide.addShape(p.shapes.RECTANGLE, { x:x+0.02, y:y+h-0.42, w:w-0.04, h:0.42, fill:{ color:C.darker, transparency:22 } });
  slide.addText(caption, {
    x:x+0.18, y:y+h-0.40, w:w-0.36, h:0.38, margin:0, valign:"middle",
    fontFace:F, fontSize:9.5, italic:true, color:C.white,
  });
}

// Điểm gạch đầu dòng (ô vuông vàng đồng tự vẽ)
function bullet(slide, x, y, w, text, opts = {}){
  slide.addShape(p.shapes.ROUNDED_RECTANGLE, {
    x:x, y:y+0.09, w:0.13, h:0.13, rectRadius:0.03,
    fill:{ color: opts.sq || C.gold },
  });
  slide.addText(text, {
    x:x+0.30, y:y-0.04, w:w-0.30, h:opts.h || 0.95, margin:0, valign:"top",
    fontFace:F, fontSize:opts.size || 14, color:opts.color || C.ink,
    lineSpacingMultiple:1.08,
  });
}

// Thẻ nhấn nổi bật
function chip(slide, x, y, w, h, head, body, gold){
  slide.addShape(p.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius:0.10,
    fill:{ color: gold ? C.gold : C.dark },
    shadow: sh(),
  });
  if (head){
    slide.addText(head, {
      x:x+0.30, y:y+0.14, w:w-0.6, h:0.34, margin:0,
      fontFace:F, fontSize:12.5, bold:true, color: gold ? C.dark : C.goldLt, charSpacing:1,
    });
  }
  slide.addText(body, {
    x:x+0.30, y:y+(head?0.52:0.16), w:w-0.6, h:h-(head?0.62:0.30), margin:0, valign:"middle",
    fontFace:F, fontSize: gold ? 12.8 : 12.5, bold:gold, italic:!gold,
    color: gold ? C.darker : C.white, lineSpacingMultiple:1.1,
  });
}

/* ============ SLIDE 1 — BÌA ============ */
let s = p.addSlide();
s.addImage({ path: IMG + "halong_hero.jpg", x:0, y:0, w:W, h:H });
s.addShape(p.shapes.RECTANGLE, { x:0, y:0, w:W, h:H, fill:{ color:C.darker, transparency:42 } });
s.addShape(p.shapes.RECTANGLE, { x:0, y:4.35, w:W, h:3.15, fill:{ color:C.darker, transparency:18 } });
s.addShape(p.shapes.RECTANGLE, { x:0.95, y:2.02, w:1.35, h:0.06, fill:{ color:C.gold } });
s.addText("TỈNH QUẢNG NINH  ·  MIỀN ĐÔNG BẮC TỔ QUỐC", {
  x:0.95, y:2.22, w:11, h:0.4, margin:0,
  fontFace:F, fontSize:15, bold:true, color:C.goldLt, charSpacing:3,
});
s.addText("LỊCH SỬ TỈNH\nQUẢNG NINH", {
  x:0.92, y:2.68, w:11.4, h:2.0, margin:0,
  fontFace:F, fontSize:50, bold:true, color:C.white, lineSpacingMultiple:0.98,
});
s.addText("Chặng đường hàng nghìn năm gắn liền với công cuộc dựng nước và giữ nước của dân tộc Việt Nam", {
  x:0.95, y:4.85, w:9.8, h:0.7, margin:0,
  fontFace:F, fontSize:16, italic:true, color:C.cream, lineSpacingMultiple:1.15,
});
s.addText("Từ các hang động tiền sử trên Vịnh Hạ Long đến tỉnh năng động hôm nay", {
  x:0.95, y:6.55, w:11, h:0.35, margin:0,
  fontFace:F, fontSize:11.5, color:C.goldLt, charSpacing:1,
});
s.addNotes("Mở đầu: giới thiệu tổng quan — Quảng Ninh là tỉnh vùng Đông Bắc, có lịch sử lâu dài gắn với dựng nước và giữ nước.");

/* ============ SLIDE 2 — MỤC LỤC / LỘ TRÌNH ============ */
s = p.addSlide();
fullBg(s, C.cream);
s.addShape(p.shapes.OVAL, { x:10.6, y:-2.2, w:5.2, h:5.2, fill:{ color:C.cream2 } });
s.addShape(p.shapes.OVAL, { x:-1.6, y:5.6, w:3.6, h:3.6, fill:{ color:C.cream2 } });
s.addText("MỤC LỤC", { x:0.55, y:0.42, w:6, h:0.32, margin:0, fontFace:F, fontSize:12.5, bold:true, color:C.jade, charSpacing:3 });
s.addText("Lộ trình lịch sử", { x:0.55, y:0.74, w:9, h:0.7, margin:0, fontFace:F, fontSize:32, bold:true, color:C.dark });
s.addShape(p.shapes.RECTANGLE, { x:0.57, y:1.50, w:1.1, h:0.05, fill:{ color:C.gold } });

const toc = [
  ["01","Tiền sử & Sơ sử","Hàng nghìn năm trước – thế kỷ II TCN"],
  ["02","Thời Bắc thuộc","Thế kỷ II TCN – thế kỷ X"],
  ["03","Phong kiến độc lập","Thế kỷ X – thế kỷ XIX"],
  ["04","Thời Pháp thuộc","1883 – 1945"],
  ["05","Thời kỳ kháng chiến","1945 – 1975"],
  ["06","Tỉnh Quảng Ninh hôm nay","1963 – nay"],
];
const cw = 3.86, ch = 2.02, gx = 0.28, gy = 0.30, x0 = 0.57, y0 = 1.95;
toc.forEach((t, i) => {
  const col = i % 3, row = Math.floor(i / 3);
  const x = x0 + col*(cw+gx), y = y0 + row*(ch+gy);
  s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y, w:cw, h:ch, rectRadius:0.10, fill:{ color:C.white }, shadow: sh() });
  s.addShape(p.shapes.RECTANGLE, { x, y, w:0.10, h:ch, fill:{ color:C.gold } });
  seal(s, x+0.32, y+0.34, 0.78, t[0], { size:23 });
  s.addText(t[1], { x:x+1.32, y:y+0.36, w:cw-1.55, h:0.75, margin:0, valign:"middle", fontFace:F, fontSize:15.5, bold:true, color:C.dark, lineSpacingMultiple:1.0 });
  s.addText(t[2], { x:x+1.32, y:y+1.18, w:cw-1.55, h:0.6, margin:0, valign:"top", fontFace:F, fontSize:11.5, color:C.jade, bold:true, lineSpacingMultiple:1.1 });
});
s.addText("Sáu chặng đường — một mạch nguồn liên tục từ quá khứ đến hiện tại", {
  x:0.57, y:6.72, w:11, h:0.35, margin:0, fontFace:F, fontSize:12.5, italic:true, color:C.gray,
});
footer(s, 2);

/* ============ SLIDE 3 — CHƯƠNG 1: TIỀN SỬ & SƠ SỬ ============ */
s = p.addSlide();
fullBg(s, C.cream);
chapterHeader(s, "01", "CHƯƠNG 01  ·  GỐC TÍCH TIỀN NHÂN", "Thời tiền sử và sơ sử", "Hàng nghìn năm trước — thế kỷ II TCN");
photo(s, "tien_su.jpg", 0.55, 1.85, 6.05, 4.95, "Cư dân tiền sử trong hang động đá vôi nhìn ra vịnh (minh hoạ)");
let bx = 7.05, bw = 5.75;
bullet(s, bx, 2.05, bw,
  "Con người đã sinh sống ở vùng Quảng Ninh từ hàng nghìn năm trước, để lại dấu tích tại các hang động và đảo trên Vịnh Hạ Long.", { h:1.35 });
bullet(s, bx, 3.35, bw,
  "Khu vực này thuộc lãnh thổ các nhà nước cổ Văn Lang và Âu Lạc — nơi cư trú của người Việt cổ.", { h:1.05 });
bullet(s, bx, 4.55, bw,
  "Địa bàn cư trú trải ven biển và các đảo đá vôi: nền tảng của văn hóa cư dân biển – rừng Đông Bắc.", { h:1.0 });
chip(s, bx, 5.65, bw, 1.05, "DẤU MỐC",
  "Quảng Ninh là bộ phận của Văn Lang – Âu Lạc, những nhà nước đầu tiên của người Việt.", true);
footer(s, 3);
s.addNotes("Nhấn mạnh: dấu tích người tiền sử ở hang động, đảo trên vịnh; vùng đất thuộc Văn Lang, Âu Lạc.");

/* ============ SLIDE 4 — CHƯƠNG 2: BẮC THUỘC ============ */
s = p.addSlide();
fullBg(s, C.cream);
chapterHeader(s, "02", "CHƯƠNG 02  ·  HƠN MỘT NGHÌN NĂM BẮC THUỘC", "Thời Bắc thuộc", "Thế kỷ II TCN — thế kỷ X");
photo(s, "hai_ba_trung.jpg", 0.55, 1.85, 6.05, 4.95, "Hưởng ứng khởi nghĩa Hai Bà Trưng (40–43 SCN) (minh hoạ)");
bullet(s, bx, 2.05, bw,
  "Quảng Ninh nằm dưới sự cai trị của các triều đại phong kiến phương Bắc suốt hơn 1.000 năm.", { h:1.0 });
bullet(s, bx, 3.10, bw,
  "Nhân dân địa phương nhiều lần đứng lên đấu tranh, tham gia các cuộc khởi nghĩa chống Bắc thuộc.", { h:1.0 });
bullet(s, bx, 4.15, bw,
  "Tiêu biểu là hưởng ứng khởi nghĩa Hai Bà Trưng và các phong trào đấu tranh liên tục khác trên vùng biên cương.", { h:1.2 });
chip(s, bx, 5.65, bw, 1.05, "TRUYỀN THỐNG",
  "Mảnh đất biên giới từ sớm đã hun đúc tinh thần bất khuất, không cam chịu ách đô hộ.", false);
footer(s, 4);
s.addNotes("Bắc thuộc: nhân dân tham gia khởi nghĩa Hai Bà Trưng và nhiều phong trào khác.");

/* ============ SLIDE 5 — CHƯƠNG 3: PHONG KIẾN ĐỘC LẬP ============ */
s = p.addSlide();
fullBg(s, C.cream);
chapterHeader(s, "03", "CHƯƠNG 03  ·  BIÊN CƯƠNG TỎA SÁNG", "Thời phong kiến độc lập", "Thế kỷ X — thế kỷ XIX");
s.addText("Vùng biên giới quan trọng: vừa phát triển kinh tế, vừa giữ vai trò phòng thủ.", {
  x:0.55, y:1.60, w:12.2, h:0.35, margin:0, fontFace:F, fontSize:12.5, italic:true, color:C.gray,
});
photo(s, "bach_dang.jpg", 0.55, 2.05, 5.75, 2.62, "Trận Bạch Đằng — bãi cọc ngăn thuyền địch (minh hoạ)");
photo(s, "yen_tu.jpg",   0.55, 4.85, 5.75, 1.95, "Chùa Đồng trên đỉnh Yên Tử, trung tâm Thiền phái Trúc Lâm");

const events = [
  ["938", "NGÔ QUYỀN ĐẠI PHÁ NAM HÁN",
   "Trên sông Bạch Đằng, Ngô Quyền đánh bại quân Nam Hán, mở ra thời kỳ độc lập lâu dài của dân tộc."],
  ["1288", "TRẦN HƯNG ĐẠO PHÁ NGUYÊN – MÔNG",
   "Trần Hưng Đạo chỉ huy quân dân Đại Việt làm nên chiến thắng Bạch Đằng vang dội, tiêu diệt đoàn thuyền quân Nguyên."],
  ["YÊN TỬ", "ĐẠO PHẬT TRÚC LÂM",
   "Yên Tử trở thành trung tâm Thiền phái Trúc Lâm do Trần Nhân Tông sáng lập — ngọn núi tổ của Phật giáo Việt Nam."],
];
const ex = 6.65, ew = 6.13, eh = 1.50, egy = 0.16;
events.forEach((e, i) => {
  const y = 2.05 + i*(eh+egy);
  s.addShape(p.shapes.ROUNDED_RECTANGLE, { x:ex, y, w:ew, h:eh, rectRadius:0.09, fill:{ color:C.white }, shadow: sh() });
  seal(s, ex+0.22, y+eh/2-0.42, 0.84, e[0], { size: e[0].length > 2 ? 15 : 22 });
  s.addText(e[1], { x:ex+1.28, y:y+0.16, w:ew-1.5, h:0.32, margin:0, fontFace:F, fontSize:12.5, bold:true, color:C.jade, charSpacing:0.5 });
  s.addText(e[2], { x:ex+1.28, y:y+0.52, w:ew-1.5, h:0.85, margin:0, valign:"top", fontFace:F, fontSize:11.8, color:C.ink, lineSpacingMultiple:1.05 });
});
footer(s, 5);
s.addNotes("Ba dấu mốc lớn: 938 Ngô Quyền, 1288 Trần Hưng Đạo, và Yên Tử – Trúc Lâm của Trần Nhân Tông.");

/* ============ SLIDE 6 — CHƯƠNG 4: PHÁP THUỘC ============ */
s = p.addSlide();
fullBg(s, C.cream);
chapterHeader(s, "04", "CHƯƠNG 04  ·  LÒ LỬA CÁCH MẠNG", "Thời Pháp thuộc", "1883 — 1945");
photo(s, "mo_than.jpg", 0.55, 1.85, 6.05, 4.95, "Thợ mỏ người Việt tại khu mỏ than thời thuộc địa (tư liệu minh hoạ)");
bullet(s, bx, 2.00, bw,
  "Thực dân Pháp ra sức khai thác các mỏ than lớn ở Hòn Gai, Cẩm Phả và Uông Bí.", { h:1.0 });
bullet(s, bx, 3.05, bw,
  "Giai cấp công nhân mỏ hình thành, nhanh chóng trở thành lực lượng cách mạng quan trọng.", { h:1.0 });
bullet(s, bx, 4.10, bw,
  "Năm 1936, nổ ra cuộc Tổng bãi công của thợ mỏ — tiếng vang lớn trong phong trào công nhân cả nước.", { h:1.1 });
chip(s, bx, 5.55, bw, 1.15, "“KỶ LUẬT VÀ ĐỒNG TÂM”",
  "Truyền thống quý báu của thợ mỏ Quảng Ninh, được hun đúc từ những cuộc đấu tranh dưới hầm lò.", true);
footer(s, 6);
s.addNotes("Pháp khai thác than; công nhân mỏ thành lực lượng cách mạng; 1936 tổng bãi công; truyền thống Kỷ luật và Đồng tâm.");

/* ============ SLIDE 7 — CHƯƠNG 5: KHÁNG CHIẾN ============ */
s = p.addSlide();
fullBg(s, C.cream);
chapterHeader(s, "05", "CHƯƠNG 05  ·  VỪA SẢN XUẤT, VỪA CHIẾN ĐẤU", "Thời kỳ kháng chiến", "1945 — 1975");
photo(s, "khang_chien.jpg", 0.55, 1.85, 6.05, 4.95, "Công nhân mỏ hăng say sản xuất phục vụ kháng chiến (minh hoạ)");
bullet(s, bx, 2.05, bw,
  "Sau Cách mạng Tháng Tám năm 1945, Quảng Ninh tham gia kháng chiến chống thực dân Pháp, sau đó là chống Mỹ cứu nước.", { h:1.35 });
bullet(s, bx, 3.45, bw,
  "Quân và dân vùng mỏ vừa giữ đất, vừa duy trì sản xuất trong điều kiện chiến tranh ác liệt.", { h:1.0 });
bullet(s, bx, 4.55, bw,
  "Than Quảng Ninh đóng vai trò quan trọng trong phát triển kinh tế và phục vụ đất nước ở hậu phương.", { h:1.0 });
chip(s, bx, 5.80, bw, 0.90, "KHẨU HIỆU",
  "“Tay búa, tay súng” — than từ Quảng Ninh theo đường vận tải ra tiền tuyến.", false);
footer(s, 7);
s.addNotes("Sau 1945: kháng chiến chống Pháp rồi chống Mỹ; than phục vụ kinh tế và chiến đấu.");

/* ============ SLIDE 8 — CHƯƠNG 6: 1963 – NAY ============ */
s = p.addSlide();
fullBg(s, C.cream);
chapterHeader(s, "06", "CHƯƠNG 06  ·  TỈNH QUẢNG NINH THỜI HỘI NHẬP", "Từ năm 1963 đến nay", "30/10/1963 — hiện tại");

// Thẻ thành lập tỉnh
s.addShape(p.shapes.ROUNDED_RECTANGLE, { x:0.55, y:1.85, w:7.55, h:1.12, rectRadius:0.10, fill:{ color:C.gold }, shadow: sh() });
seal(s, 0.80, 2.10, 0.62, "1963", { size:15 });
s.addText("NGÀY 30/10/1963 — THÀNH LẬP TỈNH QUẢNG NINH", {
  x:1.62, y:2.02, w:6.35, h:0.32, margin:0, fontFace:F, fontSize:12.5, bold:true, color:C.dark, charSpacing:0.5,
});
s.addText("Trên cơ sở hợp nhất Khu Hồng Quảng và tỉnh Hải Ninh thành một đơn vị hành chính mới.", {
  x:1.62, y:2.36, w:6.35, h:0.5, margin:0, fontFace:F, fontSize:12, color:C.darker, lineSpacingMultiple:1.05,
});

s.addText("SAU CÔNG CUỘC ĐỔI MỚI (1986) — BỐN TRỤ ĐỨT PHÁ", {
  x:0.55, y:3.20, w:7.5, h:0.3, margin:0, fontFace:F, fontSize:12, bold:true, color:C.jade, charSpacing:1,
});
const pillars = [
  ["Khai thác & chế biến than", "Ngành công nghiệp trụ cột, tiếp nối truyền thống vùng than lớn nhất cả nước."],
  ["Du lịch — Vịnh Hạ Long", "Di sản Thiên nhiên Thế giới, trung tâm du lịch biển hàng đầu Việt Nam."],
  ["Thương mại biên giới", "Cửa khẩu quốc tế Móng Cái, giao thương sôi động với Trung Quốc."],
  ["Công nghiệp & hạ tầng", "Công nghiệp, dịch vụ hiện đại; cảng biển, giao thông kết nối đồng bộ."],
];
const pw = 3.68, ph = 1.42, pgx = 0.19, pgy = 0.18;
pillars.forEach((it, i) => {
  const col = i % 2, row = Math.floor(i/2);
  const x = 0.55 + col*(pw+pgx), y = 3.60 + row*(ph+pgy);
  s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y, w:pw, h:ph, rectRadius:0.09, fill:{ color:C.white }, shadow: sh() });
  s.addShape(p.shapes.OVAL, { x:x+0.22, y:y+0.22, w:0.42, h:0.42, fill:{ color:C.jade } });
  s.addText(String(i+1), { x:x+0.22, y:y+0.22, w:0.42, h:0.42, margin:0, align:"center", valign:"middle", fontFace:F, fontSize:14, bold:true, color:C.goldLt });
  s.addText(it[0], { x:x+0.78, y:y+0.18, w:pw-0.95, h:0.5, margin:0, valign:"middle", fontFace:F, fontSize:12.5, bold:true, color:C.dark, lineSpacingMultiple:0.95 });
  s.addText(it[1], { x:x+0.24, y:y+0.72, w:pw-0.45, h:0.6, margin:0, valign:"top", fontFace:F, fontSize:10.8, color:C.gray, lineSpacingMultiple:1.03 });
});

photo(s, "quang_ninh_nay.jpg", 8.42, 1.85, 4.36, 2.78, "Thành phố Hạ Long về đêm — diện mạo tỉnh năng động");
photo(s, "mong_cai.jpg",       8.42, 4.83, 4.36, 1.97, "Cửa khẩu quốc tế Móng Cái — đầu mối thương mại biên giới");
footer(s, 8);
s.addNotes("30/10/1963 hợp nhất Hồng Quảng + Hải Ninh; sau Đổi mới: than, du lịch Hạ Long, Móng Cái, công nghiệp-dịch vụ-hạ tầng.");

/* ============ SLIDE 9 — DÒNG THỜI GIAN ============ */
s = p.addSlide();
fullBg(s, C.cream);
s.addText("TỔNG QUAN", { x:0.55, y:0.42, w:6, h:0.32, margin:0, fontFace:F, fontSize:12.5, bold:true, color:C.jade, charSpacing:3 });
s.addText("Dòng thời gian lịch sử Quảng Ninh", { x:0.55, y:0.74, w:11, h:0.7, margin:0, fontFace:F, fontSize:32, bold:true, color:C.dark });
s.addShape(p.shapes.RECTANGLE, { x:0.57, y:1.50, w:1.1, h:0.05, fill:{ color:C.gold } });

const lineY = 4.12;
s.addShape(p.shapes.RECTANGLE, { x:0.9, y:lineY+0.135, w:11.55, h:0.045, fill:{ color:C.gold } });
const milestones = [
  ["TIỀN SỬ", "Cư dân sinh sống ở hang động, đảo trên vịnh; vùng đất thuộc Văn Lang – Âu Lạc."],
  ["938", "Ngô Quyền đại phá quân Nam Hán trên sông Bạch Đằng."],
  ["1288", "Trần Hưng Đạo đánh tan quân Nguyên – Mông, trận Bạch Đằng vang dội."],
  ["1883–1945", "Pháp thuộc khai thác than Hòn Gai, Cẩm Phả, Uông Bí; công nhân mỏ hình thành."],
  ["1936", "Tổng bãi công thợ mỏ — truyền thống “Kỷ luật và Đồng tâm”."],
  ["30/10/1963", "Thành lập tỉnh Quảng Ninh trên cơ sở Hồng Quảng và Hải Ninh."],
  ["1986 – NAY", "Đổi mới: than, du lịch Hạ Long (UNESCO 1994), cửa khẩu Móng Cái, dịch vụ bứt phá."],
];
const n = milestones.length, xStart = 1.25, xStep = (12.15 - xStart)/(n-1);
milestones.forEach((m, i) => {
  const cx = xStart + i*xStep;
  const above = i % 2 === 0;
  const cwT = 1.78, chT = 1.82;
  const cy = above ? lineY - 0.32 - chT : lineY + 0.55;
  // đường nối từ trục tới thẻ
  s.addShape(p.shapes.RECTANGLE, { x:cx-0.012, y: above ? cy+chT : lineY+0.18, w:0.024, h: above ? (lineY-cy-chT+0.16) : (cy-lineY-0.40), fill:{ color:C.goldLt } });
  s.addShape(p.shapes.ROUNDED_RECTANGLE, { x:cx-cwT/2, y:cy, w:cwT, h:chT, rectRadius:0.08, fill:{ color:C.white }, shadow: sh() });
  s.addText(m[0], { x:cx-cwT/2+0.10, y:cy+0.12, w:cwT-0.20, h:0.4, margin:0, align:"center", fontFace:F, fontSize:12.5, bold:true, color:C.jade });
  s.addText(m[1], { x:cx-cwT/2+0.12, y:cy+0.52, w:cwT-0.24, h:chT-0.62, margin:0, valign:"top", align:"center", fontFace:F, fontSize:9.6, color:C.ink, lineSpacingMultiple:1.04 });
  // nút trên trục
  s.addShape(p.shapes.OVAL, { x:cx-0.19, y:lineY, w:0.38, h:0.38, fill:{ color:C.gold } });
  s.addShape(p.shapes.OVAL, { x:cx-0.10, y:lineY+0.09, w:0.20, h:0.20, fill:{ color:C.dark } });
});
footer(s, 9);

/* ============ SLIDE 10 — Ý NGHĨA LỊCH SỬ ============ */
s = p.addSlide();
fullBg(s, C.dark);
s.addImage({ path: IMG + "halong_hero.jpg", x:7.1, y:0, w:6.25, h:H, transparency:62 });
s.addShape(p.shapes.RECTANGLE, { x:7.1, y:0, w:6.25, h:H, fill:{ color:C.darker, transparency:35 } });
s.addShape(p.shapes.RECTANGLE, { x:0, y:0, w:7.6, h:H, fill:{ color:C.dark } });
s.addShape(p.shapes.RECTANGLE, { x:0.55, y:0.72, w:1.25, h:0.06, fill:{ color:C.gold } });
s.addText("Ý NGHĨA LỊCH SỬ", { x:0.55, y:0.92, w:8, h:0.35, margin:0, fontFace:F, fontSize:13, bold:true, color:C.goldLt, charSpacing:3 });
s.addText("Một vùng đất, bốn giá trị cốt lõi", { x:0.55, y:1.30, w:8, h:0.7, margin:0, fontFace:F, fontSize:28, bold:true, color:C.white });

const values = [
  ["VỊ THẾ CHIẾN LƯỢC", "Vùng biên cương Đông Bắc, cửa ngõ biển – biên của Tổ quốc; vai trò quốc phòng trọng yếu từ thời dựng nước đến nay."],
  ["CÁI NÔI NGÀNH THAN VIỆT NAM", "Vùng than lớn nhất nước — Hòn Gai, Cẩm Phả, Uông Bí — gắn với giai cấp công nhân mỏ và truyền thống đấu tranh."],
  ["DI SẢN THIÊN NHIÊN & VĂN HÓA", "Vịnh Hạ Long, non thiêng Yên Tử, chiến tích Bạch Đằng và tinh thần “Kỷ luật và Đồng tâm”."],
  ["ĐỘNG LỰC PHÁT TRIỂN", "Kinh tế biển, du lịch, thương mại biên giới và công nghiệp hiện đại — đóng góp quan trọng cho đất nước."],
];
values.forEach((v, i) => {
  const y = 2.30 + i*1.13;
  s.addShape(p.shapes.ROUNDED_RECTANGLE, { x:0.55, y, w:6.45, h:1.00, rectRadius:0.08, fill:{ color:C.deep }, line:{ color:C.gold, width:0.75 } });
  s.addShape(p.shapes.OVAL, { x:0.78, y:y+0.28, w:0.44, h:0.44, fill:{ color:C.gold } });
  s.addText(String(i+1), { x:0.78, y:y+0.28, w:0.44, h:0.44, margin:0, align:"center", valign:"middle", fontFace:F, fontSize:15, bold:true, color:C.dark });
  s.addText(v[0], { x:1.42, y:y+0.13, w:5.4, h:0.32, margin:0, fontFace:F, fontSize:12.5, bold:true, color:C.goldLt, charSpacing:0.5 });
  s.addText(v[1], { x:1.42, y:y+0.46, w:5.45, h:0.48, margin:0, valign:"top", fontFace:F, fontSize:10.6, color:C.cream, lineSpacingMultiple:1.03 });
});
s.addText("Quảng Ninh góp phần quan trọng vào sự phát triển kinh tế và lịch sử của đất nước.", {
  x:7.55, y:5.55, w:5.2, h:1.4, margin:0, fontFace:F, fontSize:15, italic:true, bold:true, color:C.white, lineSpacingMultiple:1.2, align:"center", valign:"middle",
});
footer(s, 10);

/* ============ SLIDE 11 — KẾT ============ */
s = p.addSlide();
s.addImage({ path: IMG + "quang_ninh_nay.jpg", x:0, y:0, w:W, h:H });
s.addShape(p.shapes.RECTANGLE, { x:0, y:0, w:W, h:H, fill:{ color:C.darker, transparency:35 } });
s.addShape(p.shapes.RECTANGLE, { x:0, y:2.35, w:W, h:2.85, fill:{ color:C.darker, transparency:22 } });
s.addShape(p.shapes.RECTANGLE, { x:5.99, y:2.75, w:1.35, h:0.06, fill:{ color:C.gold } });
s.addText("QUẢNG NINH", {
  x:1, y:2.95, w:11.33, h:1.0, margin:0, align:"center",
  fontFace:F, fontSize:52, bold:true, color:C.white, charSpacing:6,
});
s.addText("Nơi thiên nhiên, lịch sử và con người cùng toả sáng", {
  x:1, y:4.05, w:11.33, h:0.5, margin:0, align:"center",
  fontFace:F, fontSize:18, italic:true, color:C.goldLt,
});
s.addText("CẢM ƠN ĐÃ LẮNG NGHE", {
  x:1, y:5.05, w:11.33, h:0.4, margin:0, align:"center",
  fontFace:F, fontSize:13, bold:true, color:C.cream, charSpacing:4,
});

p.writeFile({ fileName: "Lich_su_Quang_Ninh.pptx" }).then(fn => {
  console.log("Đã tạo:", fn);
});
