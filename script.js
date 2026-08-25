
var DATA = {"icons": {"Витамины": "ti-pill", "Витамин D": "ti-sun-high", "Мультивитамины": "ti-grid-dots", "Минералы": "ti-diamond", "Железо": "ti-droplet", "Омега-3": "ti-fish", "Аминокислоты": "ti-atom-2", "Спортивное питание": "ti-barbell", "Детские витамины": "ti-mood-kid", "Коллаген": "ti-sparkles", "Пробиотики": "ti-bacteria", "Травы и экстракты": "ti-plant-2", "Красота: волосы, кожа, ногти": "ti-flower", "Нервная система и сон": "ti-moon-stars", "Пищеварение и печень": "ti-flame", "Иммунитет": "ti-shield-check", "Антиоксиданты": "ti-shield-half-filled", "Эфирные масла": "ti-droplet-half-2", "Другое": "ti-package"}, "hues": {"Витамины": "#C2541F", "Витамин D": "#CE8A16", "Мультивитамины": "#A65C2C", "Минералы": "#5F7386", "Железо": "#9E3B2E", "Омега-3": "#2C7D6C", "Аминокислоты": "#75689F", "Спортивное питание": "#3E4A57", "Детские витамины": "#DB7A3C", "Коллаген": "#B26079", "Пробиотики": "#5A8A55", "Травы и экстракты": "#6C8A3C", "Красота: волосы, кожа, ногти": "#BA6489", "Нервная система и сон": "#4F5B86", "Пищеварение и печень": "#B0752A", "Иммунитет": "#3B7E86", "Антиоксиданты": "#87683D", "Эфирные масла": "#7A9046", "Другое": "#7A7873"}, "short": {"Красота: волосы, кожа, ногти": "Красота", "Нервная система и сон": "Нервная система", "Пищеварение и печень": "Пищеварение", "Травы и экстракты": "Травы"}};
/* Оставьте пустым — адрес определится сам по месту размещения.
   Для боевого домена впишите его сюда: "https://catalog.yoshlik.uz" */
var SITE = "";
function baseURL(){
  if(SITE) return SITE.replace(/\/$/,'');
  return (location.origin + location.pathname)
    .replace(/index\.html$/i,'').replace(/\/$/,'');
}

/* ---------- источник данных ----------
   Файл data.xlsx лежит рядом с index.html.
   Чтобы обновить каталог — просто замените этот файл. */
var DATA_FILE = "data.xlsx";
var DATA_TAB  = "Продукты";
var P=[], CATS=[], BRANDS=[], SHELF_A=[], SHELF_B=[];
var ICONS=DATA.icons, HUES=DATA.hues, SHORT=DATA.short;
var CAT_UZ={}, FORM_UZ={}, UNIT_UZ={}, COUNTRY_UZ={};
var READY=false;
var main=document.getElementById('main'), topbar=document.getElementById('top');
var pbar=document.getElementById('pbar');
var PAGE=24, current=null, LANG='ru';

/* ---------------- i18n ---------------- */
var T={
ru:{
 nav_cat:'Каталог',
 hero_eyebrow:'каталог продукции',
 hero_l1:'Всё о продукции —', hero_l2_pre:'в ', hero_l2_em:'одной ссылке',
 hero_sub:'Составы, дозировки, способ применения и противопоказания. Открывается за две секунды прямо у прилавка.',
 hs_ph:'Найдите продукт — магний, D3, коллаген…', hs_find:'Найти', hs_cat:'Каталог',
 hs_often:'Часто ищут:', hs_none:'Ничего не нашлось', hs_none_sub:'Попробуйте другое слово',
 tips:['витамин D','магний','омега','коллаген'],
 sec_parts:'Разделы', sec_parts_h:'Найдите нужное по категории',
 sec_parts_p:'Каждый раздел открывает отфильтрованный каталог с поиском',
 sec_how:'Как это работает', sec_how_h:'Визит представителя продолжается после его ухода',
 s1h:'Разговор в аптеке', s1p:'Медицинский представитель рассказывает фармацевту о продукте.',
 s2h:'Один QR-код', s2p:'Фармацевт наводит камеру и получает страницу товара себе в телефон.',
 s3h:'Справочник под рукой', s3p:'Состав и дозировка остаются доступными в тот момент, когда нужны.',
 sec_port:'Портфель', sec_brands:'Бренды',
 cta_h:'Готовы посмотреть ассортимент?',
 cta_p:'Полный каталог с поиском по названию, бренду и действующему веществу.',
 cta_btn:'Открыть каталог',
 cat_h1:'Каталог продукции', cat_sub:'Поиск по названию, бренду или действующему веществу',
 q_ph:'Магний, omega, D3, коллаген…', all:'Все',
 sort_default:'По каталогу', sort_name:'По названию', sort_brand:'По бренду', sort_pack:'По размеру упаковки',
 v_grid:'Плитка', v_list:'Список', show_more:'Показать ещё',
 empty_h:'Ничего не нашлось', empty_p:'Попробуйте короче — «магний» вместо «цитрат магния 400»',
 nf_h:'Такого продукта нет', nf_p:'Возможно, ссылка устарела.', open_cat:'Открыть каталог',
 l_dose:'Дозировка', l_pack:'В упаковке', l_form:'Форма', l_country:'Страна',
 l_brand:'Бренд', l_manuf:'Производитель',
 b_desc:'Описание', b_comp:'Состав', b_use:'Способ применения',
 b_warn:'Предупреждения', b_manuf:'Производитель',
 facts_t:'Состав в порции', facts_s:'Supplement facts', copy:'Копировать',
 sim_h:'Похожие продукты', sim_s:'Из раздела «%s»',
 bm_h:'Другие продукты %s', bm_s:'Того же производителя', all_cat:'Весь каталог',
 qr_btn:'QR-код', qr_copy:'Копировать', qr_close:'Закрыть', print:'Печать',
 print_card:'Печать карточки', share:'Поделиться ссылкой',
 qp_h:'QR-коды для печати',
 qp_p:'Распечатайте лист и оставьте фармацевту — ссылка будет работать и после визита',
 t_ok:'Скопировано', t_fail:'Не удалось скопировать', totop:'Наверх',
 f_cat:'Каталог', f_all:'Все продукты', f_con:'Контакты',
 f_phone:'Телефон — добавить', f_qr:'QR-коды для печати',
 f_disc:'Не является лекарственным средством. Биологически активная добавка к пище.',
 uz_note:'', clear:'Очистить поиск', search_aria:'Поиск по каталогу', sort_aria:'Сортировка',
 w_caps:'капс.', w_tabs:'табл.', w_serv:'порций'
},
uz:{
 nav_cat:'Katalog',
 hero_eyebrow:'Rasmiy mahsulot katalogi',
 hero_l1:'Mahsulot haqida hammasi —', hero_l2_pre:'', hero_l2_em:'bitta havolada',
 hero_sub:'Tarkibi, dozasi, qoʻllash usuli va ogohlantirishlar. Dorixona peshtaxtasida ikki soniyada ochiladi.',
 hs_ph:'Mahsulotni toping — magniy, D3, kollagen…', hs_find:'Topish', hs_cat:'Katalog',
 hs_often:'Koʻp qidiriladi:', hs_none:'Hech narsa topilmadi', hs_none_sub:'Boshqa soʻzni sinab koʻring',
 tips:['D vitamini','magniy','omega','kollagen'],
 sec_parts:'Boʻlimlar', sec_parts_h:'Kerakli mahsulotni toifa boʻyicha toping',
 sec_parts_p:'Har bir boʻlim qidiruvli filtrlangan katalogni ochadi',
 sec_how:'Bu qanday ishlaydi', sec_how_h:'Vakil ketganidan keyin ham tashrif davom etadi',
 s1h:'Dorixonadagi suhbat', s1p:'Tibbiy vakil farmatsevtga mahsulot haqida soʻzlab beradi.',
 s2h:'Bitta QR-kod', s2p:'Farmatsevt kamerani qaratadi va mahsulot sahifasini telefoniga oladi.',
 s3h:'Qoʻl ostidagi maʼlumotnoma', s3p:'Tarkib va doza kerak boʻlgan paytda ochiq turadi.',
 sec_port:'Portfel', sec_brands:'Brendlar', sec_brands_p:'AQSH ishlab chiqaruvchilari mahsulotlari',
 cta_h:'Assortiment bilan tanishasizmi?',
 cta_p:'Nomi, brendi va taʼsir etuvchi moddasi boʻyicha qidiruvli toʻliq katalog.',
 cta_btn:'Katalogni ochish',
 cat_h1:'Mahsulotlar katalogi', cat_sub:'Nomi, brendi yoki taʼsir etuvchi moddasi boʻyicha qidiruv',
 q_ph:'Magniy, omega, D3, kollagen…', all:'Barchasi',
 sort_default:'Katalog boʻyicha', sort_name:'Nomi boʻyicha', sort_brand:'Brendi boʻyicha', sort_pack:'Qadoq hajmi boʻyicha',
 v_grid:'Katak', v_list:'Roʻyxat', show_more:'Yana koʻrsatish',
 empty_h:'Hech narsa topilmadi', empty_p:'Qisqaroq yozing — «magniy sitrat 400» oʻrniga «magniy»',
 nf_h:'Bunday mahsulot yoʻq', nf_p:'Ehtimol, havola eskirgan.', open_cat:'Katalogni ochish',
 l_dose:'Doza', l_pack:'Qadoqda', l_form:'Shakli', l_country:'Mamlakat',
 l_brand:'Brend', l_manuf:'Ishlab chiqaruvchi',
 b_desc:'Tavsif', b_comp:'Tarkibi', b_use:'Qoʻllash usuli',
 b_warn:'Ogohlantirishlar', b_manuf:'Ishlab chiqaruvchi',
 facts_t:'Bir porsiyadagi tarkib', facts_s:'Supplement facts', copy:'Nusxalash',
 sim_h:'Oʻxshash mahsulotlar', sim_s:'«%s» boʻlimidan',
 bm_h:'%s ning boshqa mahsulotlari', bm_s:'Xuddi shu ishlab chiqaruvchidan', all_cat:'Butun katalog',
 qr_btn:'QR-kod', qr_copy:'Nusxalash', qr_close:'Yopish', print:'Chop etish',
 print_card:'Kartochkani chop etish', share:'Havolani ulashish',
 qp_h:'Chop etish uchun QR-kodlar',
 qp_p:'Varaqni chop eting va farmatsevtga qoldiring — havola tashrifdan keyin ham ishlaydi',
 t_ok:'Nusxalandi', t_fail:'Nusxalab boʻlmadi', totop:'Yuqoriga',
 f_about:'Oʻzbekistonda vitaminlar, minerallar va biologik faol qoʻshimchalar importi va distribyutsiyasi. 1991 yildan beri bir qadam oldinda.',
 f_cat:'Katalog', f_all:'Barcha mahsulotlar', f_con:'Kontaktlar',
 f_phone:'Telefon — qoʻshish', f_qr:'Chop etish uchun QR-kodlar',
 f_disc:'Dori vositasi emas. Biologik faol qoʻshimcha.',
 uz_note:'Mahsulot matnlari hozircha rus tilida',
 clear:'Qidiruvni tozalash', search_aria:'Katalog boʻyicha qidiruv', sort_aria:'Saralash',
 w_caps:'kaps.', w_tabs:'tab.', w_serv:'porsiya'
}};
function t(k){ return T[LANG][k]; }
function L(h){ return LANG==='uz' ? h+(h.indexOf('?')>-1?'&':'?')+'lang=uz' : h; }

function esc(s){ return String(s==null?'':s).replace(/[&<>"']/g,function(c){
  return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }
function norm(s){ return String(s||'').toLowerCase().replace(/ё/g,'е'); }
function catLabel(c){
  if(LANG==='uz') return String(CAT_UZ[c]||c).split(':')[0].trim();
  return SHORT[c]||c;
}
function catFull(c){ return LANG==='uz' ? (CAT_UZ[c]||c) : c; }
function formLabel(f){ return LANG==='uz' ? (FORM_UZ[f]||f) : f; }
function unitLabel(u){ return LANG==='uz' ? (UNIT_UZ[u]||u) : u; }
function countryLabel(c){ return LANG==='uz' ? (COUNTRY_UZ[c]||c) : c; }
/* поле товара: берёт *_uz, если такая колонка появится в таблице */
function F(p,k){ return (LANG==='uz' && p[k+'_uz']) ? p[k+'_uz'] : p[k]; }

function icon(c){ return ICONS[c]||'ti-package'; }
function hue(c){ return HUES[c]||'#7A7873'; }
function rgba(h,a){ var n=parseInt(h.slice(1),16);
  return 'rgba('+(n>>16&255)+','+(n>>8&255)+','+(n&255)+','+a+')'; }
function mix(h,a){ return 'linear-gradient(152deg,'+rgba(h,a)+','+rgba(h,a*0.32)+')'; }
function rxesc(s){ return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
function hl(text,terms){
  var out=esc(text);
  terms.forEach(function(x){
    if(x.length<2) return;
    out=out.replace(new RegExp('('+rxesc(x)+')','gi'),'<mark>$1</mark>');
  });
  return out;
}
function nProducts(n){
  if(LANG==='uz') return n+' mahsulot';
  var a=n%10,b=n%100,w;
  if(a===1&&b!==11) w='продукт';
  else if(a>=2&&a<=4&&(b<10||b>=20)) w='продукта';
  else w='продуктов';
  return n+' '+w;
}

var counts={};
function prepare(){
  P.forEach(function(p){
    p._idx=norm([p.name,p.brand,p.category,p.category_extra,p.short,p.composition,p.form,
                 CAT_UZ[p.category]||'',CAT_UZ[p.category_extra]||''].join(' '));
  });
  counts={};
  P.forEach(function(p){
    counts[p.category]=(counts[p.category]||0)+1;
    if(p.category_extra) counts[p.category_extra]=(counts[p.category_extra]||0)+1;
  });
  // порядок из листа «Категории», новые из товаров — в конец
  var known=CATS.slice();
  Object.keys(counts).forEach(function(c){ if(known.indexOf(c)===-1) known.push(c); });
  CATS=known;
  // бренды в порядке появления
  BRANDS=[];
  P.forEach(function(p){ if(p.brand && BRANDS.indexOf(p.brand)===-1) BRANDS.push(p.brand); });
  // витрина: по одному товару из каждой категории, затем остальные
  var seen={}; SHELF_A=[]; SHELF_B=[];
  P.forEach(function(p){
    if(!seen[p.category] && SHELF_A.length<16){ seen[p.category]=1; SHELF_A.push(p.slug); }
  });
  P.forEach(function(p){
    if(SHELF_A.indexOf(p.slug)===-1 && SHELF_B.length<16) SHELF_B.push(p.slug);
  });
}
function byCat(c){ return P.filter(function(p){ return p.category===c||p.category_extra===c; }); }
function bySlug(s){ for(var i=0;i<P.length;i++) if(P[i].slug===s) return P[i]; return null; }
function doseStr(p){ return p.dose?p.dose+(p.unit?' '+unitLabel(p.unit):''):''; }
function formWord(form){
  var f=norm(form);
  return f.indexOf('таблет')>-1?t('w_tabs'):(f.indexOf('порош')>-1?t('w_serv'):t('w_caps'));
}
function packStr(p){
  if(!p.count||p.count===1) return '';
  var f=norm(p.form);
  return p.count+' '+(f.indexOf('таблет')>-1?t('w_tabs'):(f.indexOf('порош')>-1?t('w_serv'):t('w_caps')));
}
var SEARCH_MAP={
  'витамин':'vitamin','витамины':'vitamin','магний':'magnesium','кальций':'calcium',
  'цинк':'zinc','железо':'iron','селен':'selenium','хром':'chromium','йод':'iodine',
  'омега':'omega','коллаген':'collagen','биотин':'biotin','инозитол':'inositol',
  'мелатонин':'melatonin','таурин':'taurine','аргинин':'arginine','карнитин':'carnitine',
  'цистеин':'cysteine','лизин':'lysine','глутатион':'glutathione','лецитин':'lecithin',
  'берберин':'berberine','куркума':'turmeric','эхинацея':'echinacea','спирулина':'spirulina',
  'пробиотик':'probiotic','пробиотики':'probiotic','прополис':'propolis','рыбий':'fish',
  'кислота':'acid','фолиевая':'folic','медь':'copper','марганец':'manganese','калий':'potassium'
};
function expandTerm(w){
  var out=[w];
  if(SEARCH_MAP[w]) out.push(SEARCH_MAP[w]);
  var lat={'а':'a','в':'b','е':'e','к':'k','м':'m','н':'h','о':'o','р':'p','с':'c','т':'t','у':'y','х':'x','д':'d'};
  if(w.length<=2){
    var tr=w.split('').map(function(ch){ return lat[ch]||ch; }).join('');
    if(tr!==w) out.push(tr);
  }
  return out;
}
function terms(){ var q=norm(state.q).trim(); return q?q.split(/\s+/):[]; }

/* ---------------- chrome ---------------- */
function paintChrome(){
  document.documentElement.lang=LANG;
  document.getElementById('nav-home').href=L('#/');
  var nc=document.getElementById('nav-cat');
  nc.href=L('#/catalog'); nc.querySelector('span').textContent=t('nav_cat');
  [].forEach.call(document.querySelectorAll('#lang button'),function(b){
    b.setAttribute('aria-pressed', b.dataset.l===LANG);
  });
  document.getElementById('f-about').textContent=t('f_about');
  document.getElementById('f-cat-h').textContent=t('f_cat');
  document.getElementById('f-con-h').textContent=t('f_con');
  document.getElementById('f-disc').textContent=t('f_disc');
  var fa=document.getElementById('f-all');
  fa.textContent=t('f_all'); fa.href=L('#/catalog');
  [['f-c1','Витамин D'],['f-c2','Омега-3'],['f-c3','Детские витамины']].forEach(function(x){
    var el=document.getElementById(x[0]);
    el.textContent=catFull(x[1]);
    el.href=L('#/catalog?cat='+encodeURIComponent(x[1]));
  });
  document.getElementById('f-phone').textContent=t('f_phone');
  var fq=document.getElementById('f-qr');
  fq.textContent=t('f_qr'); fq.href=L('#/qr');
  document.getElementById('totop').setAttribute('aria-label',t('totop'));
  document.getElementById('qrcopy').innerHTML='<i class="ti ti-link" aria-hidden="true"></i> '+esc(t('qr_copy'));
  document.getElementById('qrclose').innerHTML='<i class="ti ti-x" aria-hidden="true"></i> '+esc(t('qr_close'));
  document.getElementById('pbar-qr').setAttribute('aria-label',t('qr_btn'));
  document.getElementById('pbar-share').setAttribute('aria-label',t('share'));
  if(READY) markSource();
}
document.getElementById('lang').addEventListener('click',function(e){
  var b=e.target.closest('button[data-l]'); if(!b||b.dataset.l===LANG) return;
  LANG=b.dataset.l;
  var h=location.hash.replace(/[?&]lang=(ru|uz)/,'');
  location.replace(L(h||'#/'));
  paintChrome();
  if(READY) route();
});

/* ---------------- cards ---------------- */
function cardHTML(p,tm){
  tm=tm||[];
  var h=hue(p.category), bits=[];
  var d=doseStr(p); if(d) bits.push('<b>'+esc(d)+'</b>');
  var k=packStr(p); if(k) bits.push(esc(k));
  if(!bits.length) bits.push(esc(formLabel(p.form)));
  var art=p.image?'<img src="'+esc(p.image)+'" alt="'+esc(F(p,'name'))+'" loading="lazy">'
    :'<i class="ti '+icon(p.category)+' ghost" aria-hidden="true" style="color:'+rgba(h,.42)+'"></i>';
  return '<a class="card rv" href="'+L('#/p/'+esc(p.slug))+'" style="--h:'+h+'">'
    + '<div class="thumb" style="background-color:var(--cream);background-image:'+mix(h,.20)+'">'+art
    + '</div>'
    + '<div class="card-b"><span class="card-brand" style="color:'+h+'">'+hl(p.brand,tm)+'</span>'
    + '<span class="card-name">'+hl(F(p,'name'),tm)+'</span>'
    + '<span class="spec">'+bits.join('<span aria-hidden="true">·</span>')+'</span></div></a>';
}
function rowHTML(p,tm){
  var h=hue(p.category);
  var art=p.image?'<img src="'+esc(p.image)+'" alt="" loading="lazy">'
    :'<i class="ti '+icon(p.category)+'" aria-hidden="true"></i>';
  return '<a class="lrow rv" href="'+L('#/p/'+esc(p.slug))+'">'
    + '<span class="li" style="background:'+rgba(h,.13)+';color:'+h+'">'+art+'</span>'
    + '<span class="ln2"><b>'+hl(F(p,'name'),tm)+'</b><span>'+esc(catLabel(p.category))+'</span></span>'
    + '<span class="lb">'+hl(p.brand,tm)+'</span>'
    + '<span class="ld">'+esc(doseStr(p)||'—')+'</span>'
    + '<span class="lc">'+esc(packStr(p)||formLabel(p.form))+'</span>'
    + '<i class="ti ti-chevron-right lg" aria-hidden="true"></i></a>';
}

var io=null;
function observe(){
  if(io) io.disconnect();
  io=new IntersectionObserver(function(es){
    es.forEach(function(e,i){
      if(e.isIntersecting){
        setTimeout(function(){ e.target.classList.add('in'); }, Math.min(i,7)*48);
        io.unobserve(e.target);
      }
    });
  },{rootMargin:'0px 0px -8% 0px'});
  [].forEach.call(document.querySelectorAll('.rv:not(.in)'),function(el){ io.observe(el); });
}

/* ---------------- home ---------------- */
function buildHomeSlider(brands,bc){
  var picks=[];
  brands.forEach(function(x){
    var withPhoto=P.filter(function(p){ return p.brand===x && p.image; });
    picks=picks.concat(withPhoto.slice(0,4));
  });
  if(picks.length<8) picks=P.filter(function(p){ return p.image; }).slice(0,16);

  var slides=picks.map(function(p){
    var h=hue(p.category);
    return '<a class="hs-card" href="'+L('#/p/'+esc(p.slug))+'" style="--h:'+h+'">'
      + '<span class="hs-card-ph"><img src="'+esc(p.image)+'" alt="'+esc(F(p,'name'))+'" loading="lazy"></span>'
      + '<span class="hs-card-body">'
      + '<span class="hs-card-brand">'+esc(p.brand)+'</span>'
      + '<b>'+esc(F(p,'name'))+'</b>'
      + '<span class="hs-card-cat">'+esc(catLabel(p.category))+'</span>'
      + '</span></a>';
  }).join('');

  return '<section class="home-slider">'
    + '<div class="hs-head shell">'
    + '<h2>'+esc(t('sec_brands'))+'</h2>'
    + '<span class="hs-arrows">'
    + '<button class="hs-nav" id="hsPrev" aria-label="Назад"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 18 9 12 15 6"/></svg></button>'
    + '<button class="hs-nav" id="hsNext" aria-label="Вперёд"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 18 15 12 9 6"/></svg></button>'
    + '</span></div>'
    + '<div class="hs-track" id="hsTrack">'+slides+'</div>'
    + '</section>';
}
function wireHomeSlider(){
  var track=document.getElementById('hsTrack'); if(!track) return;
  var prev=document.getElementById('hsPrev'), next=document.getElementById('hsNext');
  function step(){
    var c=track.querySelector('.hs-card');
    return c ? c.offsetWidth+18 : 260;
  }
  function upd(){
    prev.disabled = track.scrollLeft < 8;
    next.disabled = track.scrollLeft > track.scrollWidth - track.clientWidth - 8;
  }
  prev.addEventListener('click',function(){ track.scrollBy({left:-step()*2,behavior:'smooth'}); });
  next.addEventListener('click',function(){ track.scrollBy({left:step()*2,behavior:'smooth'}); });
  track.addEventListener('scroll',upd,{passive:true});
  upd();
}
function buildBrandBanners(){
  var banners=[
    {file:'now-foods',brand:'NOW Foods'},
    {file:'ecovita',brand:'ECOVITA'},
    {file:'life-extension',brand:'Life Extension'},
    {file:'solaray',brand:'Solaray'},
    {file:'vita-garden',brand:'Vita Garden'}
  ];
  var slides=banners.map(function(b){
    return '<a class="bb-slide" href="'+L('#/catalog?q='+encodeURIComponent(b.brand))+'">'
      + '<img src="images/banners/banner-'+b.file+'.jpg" alt="'+esc(b.brand)+'" loading="lazy"></a>';
  }).join('');
  var dots=banners.map(function(_,i){
    return '<button class="bb-dot" data-i="'+i+'" aria-label="Слайд '+(i+1)+'"></button>';
  }).join('');
  return '<section class="brand-banners"><div class="bb-track" id="bbTrack">'+slides+'</div>'
    + '<button class="bb-nav prev" id="bbPrev" aria-label="Назад"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 18 9 12 15 6"/></svg></button>'
    + '<button class="bb-nav next" id="bbNext" aria-label="Вперёд"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 18 15 12 9 6"/></svg></button>'
    + '<div class="bb-dots" id="bbDots">'+dots+'</div></section>';
}
function wireBrandBanners(){
  var track=document.getElementById('bbTrack'); if(!track) return;
  var slides=track.children, dots=document.querySelectorAll('.bb-dot');
  var idx=0, timer=null;
  function go(i){
    idx=(i+slides.length)%slides.length;
    track.scrollTo({left:slides[idx].offsetLeft-track.offsetLeft,behavior:'smooth'});
    dots.forEach(function(d,n){ d.classList.toggle('on',n===idx); });
  }
  function next(){ go(idx+1); }
  document.getElementById('bbPrev').addEventListener('click',function(){ go(idx-1); restart(); });
  document.getElementById('bbNext').addEventListener('click',function(){ go(idx+1); restart(); });
  dots.forEach(function(d){ d.addEventListener('click',function(){ go(+d.dataset.i); restart(); }); });
  function restart(){ clearInterval(timer); timer=setInterval(next,6000); }
  track.addEventListener('mouseenter',function(){ clearInterval(timer); });
  track.addEventListener('mouseleave',restart);
  go(0); restart();
}
function viewHome(){
  topbar.classList.add('on-dark');
  var tiles=CATS.filter(function(c){ return counts[c]; }).map(function(c){
    var h=hue(c);
    return '<a class="cat-tile rv" href="'+L('#/catalog?cat='+encodeURIComponent(c))+'">'
      + '<span class="wash" style="background-image:'+mix(h,.22)+'"></span>'
      + '<i class="ti '+icon(c)+'" aria-hidden="true" style="color:'+h+'"></i>'
      + '<b>'+esc(catLabel(c))+'</b>'
      + '<span class="n">'+esc(nProducts(counts[c]))+'</span>'
      + '<i class="ti ti-arrow-up-right arw" aria-hidden="true" style="color:'+h+'"></i></a>';
  }).join('');

  function mini(slug){
    var p=bySlug(slug); if(!p) return '';
    var h=hue(p.category);
    var art=p.image?'<img src="'+esc(p.image)+'" alt="'+esc(F(p,'name'))+'" loading="lazy">'
      :'<i class="ti '+icon(p.category)+'" aria-hidden="true"></i>';
    return '<a class="mini" href="'+L('#/p/'+esc(p.slug))+'" style="--h:'+h+'">'
      + '<span class="mi" style="background:'+rgba(h,.13)+';color:'+h+'">'+art+'</span>'
      + '<span><b>'+esc(F(p,'name'))+'</b><span>'+esc(p.brand)+'</span></span></a>';
  }
  var a=SHELF_A.map(mini).join(''), b=SHELF_B.map(mini).join('');
  var bc={}; P.forEach(function(p){ bc[p.brand]=(bc[p.brand]||0)+1; });
  var ribbon=BRANDS.concat(BRANDS).map(function(x){
    return '<span>'+esc(x)+'</span><i class="ti ti-circle-filled" aria-hidden="true"></i>';
  }).join('');
  var tips=t('tips').map(function(s){
    return '<button data-s="'+esc(s)+'">'+esc(s)+'</button>'; }).join('');

  main.innerHTML =
    '<section class="hero">'
    + '<span class="hero-glow" id="glow"></span><span class="hero-glow b"></span>'
    + '<svg class="rings" id="rings" viewBox="0 0 660 660" aria-hidden="true">'
    + '<circle cx="330" cy="330" r="200"/><circle cx="330" cy="330" r="262"/><circle cx="330" cy="330" r="324"/></svg>'
    + '<div class="shell"><div class="hero-in">'
    + '<span class="eyebrow">'+esc(t('hero_eyebrow'))+'</span>'
    + '<h1><span class="ln"><span>'+esc(t('hero_l1'))+'</span></span>'
    + '<span class="ln"><span>'+esc(t('hero_l2_pre'))+'<em>'+esc(t('hero_l2_em'))+'</em></span></span></h1>'
    + '<p class="sub">'+esc(t('hero_sub'))+'</p>'
    + '<div class="hsearch"><div class="hs-row">'
    + '<div class="hs-field"><i class="ti ti-search" aria-hidden="true"></i>'
    + '<input id="hq" type="search" placeholder="'+esc(t('hs_ph'))+'" autocomplete="off" aria-label="'+esc(t('search_aria'))+'">'
    + '<button class="hs-go" id="hgo"><span>'+esc(t('hs_find'))+'</span><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12h14M13 6l6 6-6 6"/></svg></button></div>'
    + '<a class="hs-cat" href="'+L('#/catalog')+'"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="4" width="6" height="6" rx="1"/><rect x="14" y="4" width="6" height="6" rx="1"/><rect x="4" y="14" width="6" height="6" rx="1"/><rect x="14" y="14" width="6" height="6" rx="1"/></svg><span>'+esc(t('hs_cat'))+'</span></a>'
    + '</div><div class="hs-drop" id="hdrop"></div>'
    + '<div class="hs-hint"><span>'+esc(t('hs_often'))+'</span>'+tips+'</div>'
    + '</div></div></div></section>'

    + buildBrandBanners()
    + '<section class="band-cream"><div class="shell sec">'
    + '<div class="sec-head rv"><span class="tagline-lbl">'+esc(t('sec_parts'))+'</span>'
    + '<h2>'+esc(t('sec_parts_h'))+'</h2><p>'+esc(t('sec_parts_p'))+'</p></div>'
    + '<div class="cat-grid">'+tiles+'</div></div></section>'

    + '<div class="ribbon" aria-hidden="true"><div class="rt">'+ribbon+'</div></div>'


    + '<section class="cta-band"><span class="glow"></span><div class="shell cta-in">'
    + '<div><h2>'+esc(t('cta_h'))+'</h2><p>'+esc(t('cta_p'))+'</p></div>'
    + '<a class="btn btn-primary" href="'+L('#/catalog')+'">'+esc(t('cta_btn'))+' <i class="ti ti-arrow-right" aria-hidden="true"></i></a>'
    + '</div></section>';

  heroSearch();
  wireBrandBanners();
  var glow=document.getElementById('glow'), rings=document.getElementById('rings');
  if(glow && window.matchMedia('(hover:hover)').matches){
    var hero=glow.parentNode, tk=false;
    hero.addEventListener('mousemove',function(e){
      if(tk) return; tk=true;
      requestAnimationFrame(function(){
        var r=hero.getBoundingClientRect();
        glow.style.setProperty('--gx',((e.clientX-r.left)/r.width*100)+'%');
        glow.style.setProperty('--gy',((e.clientY-r.top)/r.height*100)+'%');
        tk=false;
      });
    });
  }
  window._parallax=function(y){ if(rings) rings.style.transform='translateY('+(y*0.16)+'px)'; };
  observe();
}

function heroSearch(){
  var inp=document.getElementById('hq'), drop=document.getElementById('hdrop'), sel=-1, hits=[];
  function go(q){ location.hash=L('#/catalog'+(q?'?q='+encodeURIComponent(q):'')); }
  function render(){
    var q=norm(inp.value).trim();
    if(q.length<2){ drop.classList.remove('open'); drop.innerHTML=''; hits=[]; return; }
    var tm=q.split(/\s+/);
    hits=P.filter(function(p){
      for(var i=0;i<tm.length;i++){
        var vars=expandTerm(tm[i]), hit=false;
        for(var v=0;v<vars.length;v++){
          var needle=vars[v];
          var re=needle.length<=2
            ? new RegExp('(^|[^a-zа-я0-9])'+needle.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+'([^a-zа-я0-9]|$)','i')
            : null;
          if(re ? re.test(p._idx) : p._idx.indexOf(needle)!==-1){ hit=true; break; }
        }
        if(!hit) return false;
      }
      return true;
    }).slice(0,5);
    if(!hits.length){
      drop.innerHTML='<div class="hs-item"><span class="hs-ic" style="background:var(--cream);color:var(--faint)">'
        + '<i class="ti ti-search-off" aria-hidden="true"></i></span><span><b>'+esc(t('hs_none'))+'</b>'
        + '<span>'+esc(t('hs_none_sub'))+'</span></span></div>';
    } else {
      drop.innerHTML=hits.map(function(p,i){
        var h=hue(p.category);
        return '<a class="hs-item'+(i===sel?' sel':'')+'" href="'+L('#/p/'+esc(p.slug))+'">'
          + '<span class="hs-ic" style="background:'+rgba(h,.13)+';color:'+h+'"><i class="ti '+icon(p.category)+'" aria-hidden="true"></i></span>'
          + '<span><b>'+hl(F(p,'name'),tm)+'</b><span>'+esc(p.brand)+' · '+esc(catLabel(p.category))+'</span></span></a>';
      }).join('');
    }
    drop.classList.add('open');
  }
  inp.addEventListener('input',function(){ sel=-1; render(); });
  inp.addEventListener('keydown',function(e){
    if(e.key==='Enter'){
      if(sel>=0 && hits[sel]) location.hash=L('#/p/'+hits[sel].slug);
      else go(inp.value);
    } else if(e.key==='ArrowDown'){ e.preventDefault(); sel=Math.min(sel+1,hits.length-1); render(); }
    else if(e.key==='ArrowUp'){ e.preventDefault(); sel=Math.max(sel-1,-1); render(); }
    else if(e.key==='Escape'){ drop.classList.remove('open'); }
  });
  document.getElementById('hgo').addEventListener('click',function(){ go(inp.value); });
  [].forEach.call(document.querySelectorAll('.hs-hint button'),function(b){
    b.addEventListener('click',function(){ go(b.dataset.s); });
  });
  document.addEventListener('click',function(e){
    if(!e.target.closest('.hsearch')) drop.classList.remove('open');
  });
}

/* ---------------- catalog ---------------- */
var state={q:'',cat:'',shown:PAGE,sort:'default',view:'grid'};
var lastCatalogState=null;

function viewCatalog(params){
  topbar.classList.remove('on-dark');
  var newQ=params.get('q')||'', newCat=params.get('cat')||'';
  if(newQ!==state.q || newCat!==state.cat){ state.shown=PAGE; }
  state.q=newQ; state.cat=newCat;

  var chips=['<button class="chip" data-cat="" aria-pressed="'+(!state.cat)+'">'+esc(t('all'))+' <em>'+P.length+'</em></button>']
    .concat(CATS.filter(function(c){ return counts[c]; }).map(function(c){
      return '<button class="chip" data-cat="'+esc(c)+'" aria-pressed="'+(state.cat===c)+'">'
        + esc(catLabel(c))+' <em>'+counts[c]+'</em></button>';
    })).join('');

  main.innerHTML =
    '<div class="catalog-page">'
    + '<div class="shell cat-head">'+esc(t('cat_h1'))+'</h1><p>'+esc(t('cat_sub'))+'</p></div>'
    + '<div class="filters"><div class="shell">'
    + '<div class="search-wrap"><i class="ti ti-search" aria-hidden="true"></i>'
    + '<input id="q" type="search" placeholder="'+esc(t('q_ph'))+'" value="'+esc(state.q)+'" aria-label="'+esc(t('search_aria'))+'">'
    + '<span class="kbd" id="kbd">/</span>'
    + '<button class="clear-q" id="clearq" aria-label="'+esc(t('clear'))+'"><i class="ti ti-x" aria-hidden="true"></i></button></div>'
    + '<div class="chips" id="chips">'+chips+'</div></div></div>'
    + '<div class="shell"><div class="toolbar">'
    + '<p class="count-line" id="cline"></p><span class="tools">'
    + '<select id="sort" aria-label="'+esc(t('sort_aria'))+'">'
    + '<option value="default">'+esc(t('sort_default'))+'</option>'
    + '<option value="name">'+esc(t('sort_name'))+'</option>'
    + '<option value="brand">'+esc(t('sort_brand'))+'</option>'
    + '<option value="pack">'+esc(t('sort_pack'))+'</option></select>'
    + '<span class="vtog"><button id="vgrid" aria-pressed="true" aria-label="'+esc(t('v_grid'))+'"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="4" width="6" height="6" rx="1"/><rect x="14" y="4" width="6" height="6" rx="1"/><rect x="4" y="14" width="6" height="6" rx="1"/><rect x="14" y="14" width="6" height="6" rx="1"/></svg></button>'
    + '<button id="vlist" aria-pressed="false" aria-label="'+esc(t('v_list'))+'"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><line x1="8" y1="6" x2="20" y2="6"/><line x1="8" y1="12" x2="20" y2="12"/><line x1="8" y1="18" x2="20" y2="18"/><line x1="4" y1="6" x2="4.01" y2="6"/><line x1="4" y1="12" x2="4.01" y2="12"/><line x1="4" y1="18" x2="4.01" y2="18"/></svg></button></span>'
    + '</span></div><div id="out"></div><div class="more-wrap" id="morewrap"></div></div>'
    + '</div>';
    
  var input=document.getElementById('q'), clear=document.getElementById('clearq'), kbd=document.getElementById('kbd');
  function toggleClear(){
    var on=!!input.value;
    clear.style.display=on?'grid':'none';
    kbd.style.display=on?'none':'';
  }
  toggleClear();
  input.addEventListener('input',function(){
    state.q=input.value; state.shown=PAGE; toggleClear(); renderList(); syncHash();
  });
  clear.addEventListener('click',function(){
    input.value=''; state.q=''; state.shown=PAGE; toggleClear(); renderList(); syncHash(); input.focus();
  });
  document.getElementById('sort').addEventListener('change',function(){
    state.sort=this.value; state.shown=PAGE; renderList();
  });
  var vg=document.getElementById('vgrid'), vl=document.getElementById('vlist');
  function setView(v){
    state.view=v;
    vg.setAttribute('aria-pressed',v==='grid');
    vl.setAttribute('aria-pressed',v==='list');
    renderList();
  }
  vg.addEventListener('click',function(){ setView('grid'); });
  vl.addEventListener('click',function(){ setView('list'); });

  var chipbox=document.getElementById('chips');
  chipbox.addEventListener('click',function(e){
    var b=e.target.closest('.chip'); if(!b) return;
    state.cat=b.dataset.cat; state.shown=PAGE;
    [].forEach.call(chipbox.querySelectorAll('.chip'),function(c){
      c.setAttribute('aria-pressed',c.dataset.cat===state.cat);
    });
    b.scrollIntoView({behavior:'smooth',block:'nearest',inline:'center'});
    renderList(); syncHash();
  });
  var act=chipbox.querySelector('.chip[aria-pressed="true"]');
  if(act) act.scrollIntoView({block:'nearest',inline:'center'});

  window._slash=function(){ input.focus(); input.select(); };
  renderList();
}

function syncHash(){
  var p=[];
  if(state.q) p.push('q='+encodeURIComponent(state.q));
  if(state.cat) p.push('cat='+encodeURIComponent(state.cat));
  if(LANG==='uz') p.push('lang=uz');
  history.replaceState(null,'','#/catalog'+(p.length?'?'+p.join('&'):''));
}
function filtered(){
  var tm=terms();
  var list=P.filter(function(p){
    if(state.cat && p.category!==state.cat && p.category_extra!==state.cat) return false;
    for(var i=0;i<tm.length;i++){
        var vars=expandTerm(tm[i]), hit=false;
        for(var v=0;v<vars.length;v++){
          var needle=vars[v];
          var re=needle.length<=2
            ? new RegExp('(^|[^a-zа-я0-9])'+needle.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+'([^a-zа-я0-9]|$)','i')
            : null;
          if(re ? re.test(p._idx) : p._idx.indexOf(needle)!==-1){ hit=true; break; }
        }
        if(!hit) return false;
      }
    return true;
  });
  if(state.sort==='name') list=list.slice().sort(function(a,b){ return a.name.localeCompare(b.name,'ru'); });
  else if(state.sort==='brand') list=list.slice().sort(function(a,b){
    return a.brand.localeCompare(b.brand,'ru')||a.name.localeCompare(b.name,'ru'); });
  else if(state.sort==='pack') list=list.slice().sort(function(a,b){ return (b.count||0)-(a.count||0); });
  return list;
}
function renderList(appendOnly){
  var list=filtered(), out=document.getElementById('out'),
      cline=document.getElementById('cline'), more=document.getElementById('morewrap');
  var tm=terms();
  if(!list.length){
    cline.textContent='';
    out.innerHTML='<div class="empty"><i class="ti ti-search-off" aria-hidden="true"></i>'
      + '<b>'+esc(t('empty_h'))+'</b><span>'+esc(t('empty_p'))+'</span></div>';
    more.innerHTML=''; return;
  }
  cline.textContent=nProducts(list.length);
  var page=list.slice(0,state.shown);
  var wrap=out.querySelector(state.view==='grid'?'.grid':'.list');

  if(appendOnly && wrap){
    var prevShown=state.shown-PAGE;
    var extra=list.slice(prevShown,state.shown);
    wrap.insertAdjacentHTML('beforeend', extra.map(function(p){
      return state.view==='grid' ? cardHTML(p,tm) : rowHTML(p,tm);
    }).join(''));
  } else {
    out.innerHTML = state.view==='grid'
      ? '<div class="grid">'+page.map(function(p){ return cardHTML(p,tm); }).join('')+'</div>'
      : '<div class="list">'+page.map(function(p){ return rowHTML(p,tm); }).join('')+'</div>';
  }

  more.innerHTML = list.length>state.shown
    ? '<button class="btn-outline" id="more">'+esc(t('show_more'))+' <i class="ti ti-chevron-down" aria-hidden="true"></i></button>' : '';
  if(list.length>state.shown){
    document.getElementById('more').addEventListener('click',function(){
      state.shown+=PAGE; renderList(true);
    });
  }
  observe();
}

/* ---------------- product ---------------- */
function nameStem(name){
  return norm(name)
    .replace(/\d+[.,]?\d*/g,'')
    .replace(/\b(мг|мкг|ме|iu|mcg|mg|xb|капсул\w*|таблет\w*|softgels?|vcaps?|caps?|kapsula\w*|tabletka\w*|шт\.?)\b/g,'')
    .replace(/[(),]/g,' ')
    .replace(/\s+/g,' ').trim();
}
function sameDose(p){
  var stem=nameStem(p.name);
  return P.filter(function(x){
    return x.slug!==p.slug && x.brand===p.brand && nameStem(x.name)===stem;
  });
}
function similar(p){
  var pool=byCat(p.category).filter(function(x){ return x.slug!==p.slug; });
  pool.sort(function(a,b){ return (b.brand===p.brand)-(a.brand===p.brand); });
  if(pool.length<4 && p.category_extra){
    byCat(p.category_extra).forEach(function(x){
      if(x.slug!==p.slug && pool.indexOf(x)===-1) pool.push(x);
    });
  }
  return pool.slice(0,4);
}
function sameBrand(p){
  return P.filter(function(x){ return x.brand===p.brand && x.slug!==p.slug; }).slice(0,4);
}
function factsList(text){
  var parts=String(text).split(/;\s*/).filter(Boolean);
  if(parts.length===1) parts=String(text).split(/(?:\.\s+)(?=[А-ЯA-Z])/).filter(Boolean);
  return parts.map(function(s){
    s=s.trim().replace(/[.;]$/,'');
    var m=s.match(/^(.+?)\s+[—–-]\s+([^—–]*\d[^—–]*)$/);
    if(m) return '<li><b>'+esc(m[1])+'</b><em>'+esc(m[2])+'</em></li>';
    return '<li>'+esc(s)+'</li>';
  }).join('');
}

function viewProduct(slug){
  topbar.classList.remove('on-dark');
  var p=bySlug(slug); current=p;
  if(!p){
    main.innerHTML='<div class="shell"><div class="empty"><i class="ti ti-mood-empty" aria-hidden="true"></i>'
      + '<b>'+esc(t('nf_h'))+'</b><span>'+esc(t('nf_p'))+'</span>'
      + '<p style="margin-top:20px"><a class="btn-outline" href="'+L('#/catalog')+'">'+esc(t('open_cat'))+'</a></p></div></div>';
    return;
  }
  document.title=F(p,'name');
  document.getElementById('pbar-name').textContent=F(p,'name');
  var h=hue(p.category);

  var ks=[]; var d=doseStr(p); if(d) ks.push([t('l_dose'),d]);
  var k=packStr(p); if(k) ks.push([t('l_pack'),k]);
  ks.push([t('l_form'),formLabel(p.form)]);
  if(p.country) ks.push([t('l_country'),countryLabel(p.country)]);
  var tags=[p.category].concat(p.category_extra?[p.category_extra]:[]);
  var packs=[{count:p.count,image:p.image}];
  Object.keys(p).filter(function(k){ return /^image_extra_\d+$/.test(k); })
    .map(function(k){ return parseInt(k.match(/\d+$/)[0],10); })
    .sort(function(a,b){ return a-b; })
    .forEach(function(n){
      var img=p['image_extra_'+n], cnt=p['packs_extra_'+n];
      if(img||cnt){ packs.push({count:cnt||p.count, image:img||p.image}); }
    });
  var hasMulti=packs.length>1;
  var shotSlides=packs.map(function(v){
    var im=v.image?'<img src="'+esc(v.image)+'" alt="'+esc(F(p,'name'))+'">'
      :'<i class="ti '+icon(p.category)+' ghost" aria-hidden="true" style="color:'+rgba(h,.34)+'"></i>';
    return '<div class="ps-slide">'+im+'</div>';
  }).join('');

  var secs=[['opis',t('b_desc')],['sostav',t('b_comp')],['prim',t('b_use')]];
  if(p.warnings) secs.push(['pred',t('b_warn')]);
  secs.push(['proiz',t('b_manuf')]);

  var note = (LANG==='uz' && !p.full_uz)
    ? '<p class="uz-note"><i class="ti ti-info-circle" aria-hidden="true"></i>'+esc(t('uz_note'))+'</p>' : '';

  main.innerHTML =
    '<section class="p-hero"><span class="wash" style="background-image:linear-gradient(178deg,'+rgba(h,.16)+',rgba(255,255,255,0))"></span>'
    + '<div class="shell"><div class="back-bar">'
    + '<a class="back" href="'+(lastCatalogState?lastCatalogState.hash:L('#/catalog?cat='+encodeURIComponent(p.category)))+'"><i class="ti ti-arrow-left" aria-hidden="true"></i> '+esc(catLabel(p.category))+'</a>'    + '<div class="acts">'
    + '<button class="icon-btn" id="printbtn" aria-label="'+esc(t('print_card'))+'"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg></button>'
    + '<button class="icon-btn wide" id="qrbtn"><i class="ti ti-qrcode" aria-hidden="true"></i> '+esc(t('qr_btn'))+'</button>'
    + '<button class="icon-btn" id="sharebtn" aria-label="'+esc(t('share'))+'"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><line x1="8.59" y1="13.51" x2="15.42" y2="17.49"/><line x1="15.41" y1="6.51" x2="8.59" y2="10.49"/></svg></button>'
    + '</div></div>'
    + '<div class="p-top">'
    + '<div class="p-shot" style="background-color:var(--cream);background-image:'+mix(h,.24)+'">'
    + '<div class="ps-track" id="psTrack">'+shotSlides+'</div>'
    + (hasMulti?'<button class="ps-nav prev" id="psPrev" aria-label="Предыдущее фото"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 18 9 12 15 6"/></svg></button>'
      + '<button class="ps-nav next" id="psNext" aria-label="Следующее фото"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 18 15 12 9 6"/></svg></button>'
      + '<div class="ps-dots">'+packs.map(function(_,i){return '<span class="ps-dot'+(i===0?' on':'')+'" data-i="'+i+'"></span>';}).join('')+'</div>':'')
    + '<span class="mono" style="color:'+rgba(h,.7)+'">'+esc(p.brand)+'</span></div>'
    + '<div><span class="p-brand" style="color:'+h+';background:'+rgba(h,.11)+'">'+esc(p.brand)+'</span>'
    + '<h1 class="p-title">'+esc(F(p,'name'))+'</h1>'
    + note
    + '<p class="p-lead">'+esc(F(p,'short'))+'</p>'
    + '<div class="keyspecs">'+ks.map(function(x){
        return '<div class="ks"><span>'+esc(x[0])+'</span><b>'+esc(x[1])+'</b></div>'; }).join('')+'</div>'
    + (hasMulti?'<div class="pack-pills">'+packs.map(function(v,i){
        return '<button class="pack-pill'+(i===0?' on':'')+'" data-i="'+i+'">'+esc(v.count)+' '+esc(formWord(p.form))+'</button>';
      }).join('')+'</div>':'')
    + '<div class="tagline">'+tags.map(function(x){
        var th=hue(x);
        return '<a class="tag" href="'+L('#/catalog?cat='+encodeURIComponent(x))+'" style="color:'+th+';background:'+rgba(th,.12)+'">'+esc(catLabel(x))+'</a>';
      }).join('')+'</div>'
    + (sameDose(p).length ? '<div class="dose-links"><span>Другие дозировки:</span>'
        + sameDose(p).map(function(x){
            return '<a href="'+L('#/p/'+esc(x.slug))+'">'+esc(doseStr(x)||x.name)+'</a>';
          }).join('') + '</div>' : '')
    + '</div></div></div></section>'

    + '<div class="shell"><div class="p-wrap">'
    + '<nav class="rail no-print" id="rail">'+secs.map(function(s,i){
        return '<a href="#'+s[0]+'" class="'+(i===0?'on':'')+'" data-t="'+s[0]+'">'+esc(s[1])+'</a>'; }).join('')+'</nav>'
    + '<div class="p-body">'
    + '<section class="blk rv" id="opis"><h2>'+esc(t('b_desc'))+'</h2><p>'+esc(F(p,'full'))+'</p></section>'
    + '<section class="blk rv" id="sostav"><h2>'+esc(t('b_comp'))+'</h2>'
    + '<div class="facts"><div class="facts-h"><span style="flex:1"><b>'+esc(t('facts_t'))+'</b>'
    + '<span>'+esc(t('facts_s'))+'</span></span>'
    + '<button id="copyfacts" class="no-print"><i class="ti ti-copy" aria-hidden="true"></i> '+esc(t('copy'))+'</button></div>'
    + '<ul>'+factsList(F(p,'composition'))+'</ul></div></section>'
    + '<section class="blk rv" id="prim"><h2>'+esc(t('b_use'))+'</h2><p>'+esc(F(p,'usage'))+'</p></section>'
    + (p.warnings?'<section class="blk rv" id="pred"><h2>'+esc(t('b_warn'))+'</h2>'
        + '<div class="warn"><i class="ti ti-alert-triangle" aria-hidden="true"></i><p>'+esc(F(p,'warnings'))+'</p></div></section>':'')
    + '<section class="blk rv" id="proiz"><h2>'+esc(t('b_manuf'))+'</h2><div class="meta-row">'
    + '<div><span>'+esc(t('l_brand'))+'</span><b>'+esc(p.brand)+'</b></div>'
    + (p.manufacturer?'<div><span>'+esc(t('l_manuf'))+'</span><b>'+esc(p.manufacturer)+'</b></div>':'')
    + (p.country?'<div><span>'+esc(t('l_country'))+'</span><b>'+esc(countryLabel(p.country))+'</b></div>':'')
    + '</div></section></div></div></div>'

    + '<section class="similar"><div class="shell">'
    + '<h2>'+esc(t('sim_h'))+'</h2><p class="sh">'+esc(t('sim_s').replace('%s',catFull(p.category)))+'</p>'
    + '<div class="grid">'+similar(p).map(function(x){ return cardHTML(x); }).join('')+'</div>'
    + (sameBrand(p).length ? '<h2 style="margin-top:48px">'+esc(t('bm_h').replace('%s',p.brand))+'</h2>'
        + '<p class="sh">'+esc(t('bm_s'))+'</p>'
        + '<div class="grid">'+sameBrand(p).map(function(x){ return cardHTML(x); }).join('')+'</div>' : '')
    + '<div class="more-wrap" style="padding:30px 0 0"><a class="btn-outline" href="'+L('#/catalog')+'">'+esc(t('all_cat'))+' <i class="ti ti-arrow-right" aria-hidden="true"></i></a></div>'
    + '</div></section>';
  if(hasMulti){
    var psTrack=document.getElementById('psTrack'), psSlides=psTrack.children;
    var psDots=document.querySelectorAll('.ps-dot'), psPills=document.querySelectorAll('.pack-pill');
    var curPs=0;
    function psGo(i){
      curPs=(i+psSlides.length)%psSlides.length;
      psTrack.scrollTo({left:psSlides[curPs].offsetLeft-psTrack.offsetLeft,behavior:'smooth'});
      psDots.forEach(function(d,n){ d.classList.toggle('on',n===curPs); });
      psPills.forEach(function(pl,n){ pl.classList.toggle('on',n===curPs); });
    }
    document.getElementById('psPrev').addEventListener('click',function(){ psGo(curPs-1); });
    document.getElementById('psNext').addEventListener('click',function(){ psGo(curPs+1); });
    psDots.forEach(function(d){ d.addEventListener('click',function(){ psGo(+d.dataset.i); }); });
    psPills.forEach(function(pl){ pl.addEventListener('click',function(){ psGo(+pl.dataset.i); }); });
  }
  document.getElementById('qrbtn').addEventListener('click',function(){ openQR(p); });
  document.getElementById('sharebtn').addEventListener('click',function(){ share(p); });
  document.getElementById('printbtn').addEventListener('click',function(){ window.print(); });
  document.getElementById('copyfacts').addEventListener('click',function(){
    copy(F(p,'name')+'\n\n'+t('b_comp')+': '+F(p,'composition')+'\n\n'+t('b_use')+': '+F(p,'usage'));
  });

  var rail=document.getElementById('rail');
  rail.addEventListener('click',function(e){
    var a=e.target.closest('a'); if(!a) return;
    e.preventDefault();
    var el=document.getElementById(a.dataset.t);
    if(el) el.scrollIntoView({behavior:'smooth',block:'start'});
  });
  var ro=new IntersectionObserver(function(es){
    es.forEach(function(en){
      if(en.isIntersecting){
        [].forEach.call(rail.querySelectorAll('a'),function(a){
          a.classList.toggle('on', a.dataset.t===en.target.id);
        });
      }
    });
  },{rootMargin:'-25% 0px -65% 0px'});
  secs.forEach(function(s){
    var el=document.getElementById(s[0]); if(el) ro.observe(el);
  });
  observe();
}

/* ---------------- qr sheet ---------------- */
function viewSheet(){
  topbar.classList.remove('on-dark');
  main.innerHTML='<div class="shell">'
    + '<div class="cat-head no-print"><h1>'+esc(t('qp_h'))+'</h1><p>'+esc(t('qp_p'))+'</p>'
    + '<p style="margin-top:18px"><button class="btn-outline" id="printbtn2"><i class="ti ti-printer" aria-hidden="true"></i> '+esc(t('print'))+'</button></p></div>'
    + '<div class="sheet-grid" id="sheet"></div></div>';
  document.getElementById('printbtn2').addEventListener('click',function(){ window.print(); });
  var host=document.getElementById('sheet'), i=0;
  (function chunk(){
    var frag=document.createDocumentFragment();
    for(var n=0;n<12&&i<P.length;n++,i++){
      var p=P[i], el=document.createElement('div');
      el.className='sheet-item';
      el.innerHTML='<div class="qh"></div><b>'+esc(F(p,'name'))+'</b><span>'+esc(p.brand)+'</span>';
      makeQR(el.querySelector('.qh'), qrURL(p), 116);
      frag.appendChild(el);
    }
    host.appendChild(frag);
    if(i<P.length) requestAnimationFrame(chunk);
  })();
}
function makeQR(host,text,size){
  host.innerHTML='';
  if(typeof QRCode==='undefined'){
    host.innerHTML='<span style="font-size:11px;color:#A19E97">—</span>'; return;
  }
  new QRCode(host,{text:text,width:size,height:size,
    colorDark:'#1C1916',colorLight:'#FFFFFF',correctLevel:QRCode.CorrectLevel.M});
}

/* ---------------- modal / share ---------------- */
var ovl=document.getElementById('ovl');
function qrURL(p){ return baseURL()+'/#/p/'+p.slug+(LANG==='uz'?'?lang=uz&src=qr':'?src=qr'); }
function pageURL(p){ return baseURL()+'/#/p/'+p.slug+(LANG==='uz'?'?lang=uz':''); }
function openQR(p){
  makeQR(document.getElementById('qrbox'), qrURL(p), 208);
  document.getElementById('qrname').textContent=F(p,'name');
  document.getElementById('qrurl').textContent=pageURL(p);
  ovl.classList.add('open');
  document.getElementById('qrcopy').onclick=function(){ copy(pageURL(p)); };
  document.getElementById('qrclose').focus();
}
function closeQR(){ ovl.classList.remove('open'); }
document.getElementById('qrclose').addEventListener('click',closeQR);
ovl.addEventListener('click',function(e){ if(e.target===ovl) closeQR(); });
document.getElementById('pbar-qr').addEventListener('click',function(){ if(current) openQR(current); });
document.getElementById('pbar-share').addEventListener('click',function(){ if(current) share(current); });
function toast(m){
  var el=document.getElementById('toast');
  el.textContent=m; el.classList.add('show');
  clearTimeout(el._h); el._h=setTimeout(function(){ el.classList.remove('show'); },2200);
}
function copy(text){
  if(navigator.clipboard) navigator.clipboard.writeText(text).then(
    function(){ toast(t('t_ok')); },function(){ toast(t('t_fail')); });
  else toast(t('t_fail'));
}
function share(p){
  if(navigator.share) navigator.share({title:F(p,'name'),url:pageURL(p)}).catch(function(){});
  else copy(pageURL(p));
}

/* ---------------- загрузка данных ---------------- */
var NUMS={dose:1, count:1};

function rowsToProducts(rows){
  if(!rows || !rows.length) return [];
  var head=rows[0].map(function(h){ return String(h==null?'':h).trim(); });
  var out=[];
  for(var r=1;r<rows.length;r++){
    var o={}, any=false;
    for(var c=0;c<head.length;c++){
      var key=head[c]; if(!key) continue;
      var v=(rows[r][c]==null?'':String(rows[r][c]).trim());
      if(key==='image_url') key='image';
      if(NUMS[key]) o[key] = v===''? null : (parseFloat(v.replace(',','.'))||null);
      else o[key]=v;
      if(v!=='') any=true;
    }
    if(!any || !o.slug || !o.name) continue;
    if(o.status && o.status.toLowerCase().indexOf('опублик')===-1) continue;
    if(!o.image) o.image='';
    out.push(o);
  }
  return out;
}

function readCats(rows){
  CATS=[]; CAT_UZ={};
  if(!rows || rows.length<2) return;
  for(var r=1;r<rows.length;r++){
    var ru=String(rows[r][1]==null?'':rows[r][1]).trim();
    var uz=String(rows[r][2]==null?'':rows[r][2]).trim();
    if(!ru) continue;
    if(CATS.indexOf(ru)===-1) CATS.push(ru);
    if(uz) CAT_UZ[ru]=uz;
  }
}

function readGlossary(rows){
  FORM_UZ={}; UNIT_UZ={}; COUNTRY_UZ={};
  if(!rows) return;
  var bucket=null;
  for(var r=0;r<rows.length;r++){
    var a=String(rows[r][0]==null?'':rows[r][0]).trim();
    var b=String(rows[r][1]==null?'':rows[r][1]).trim();
    if(a && !b){
      var t=a.toLowerCase();
      if(t.indexOf('форм')===0) bucket=FORM_UZ;
      else if(t.indexOf('единиц')===0) bucket=UNIT_UZ;
      else if(t.indexOf('стран')===0) bucket=COUNTRY_UZ;
      continue;
    }
    if(a && b && bucket) bucket[a]=b;
  }
}

function sheetRows(wb, name){
  var ws=wb.Sheets[name];
  return ws ? XLSX.utils.sheet_to_json(ws, {header:1, defval:'', raw:false}) : null;
}

function loadData(){
  if(typeof XLSX==='undefined') return fail('библиотека чтения таблиц не загрузилась');
  fetch(DATA_FILE+'?_='+Date.now(), {cache:'no-store'})
    .then(function(r){
      if(!r.ok) throw new Error('файл '+DATA_FILE+' не найден (HTTP '+r.status+')');
      return r.arrayBuffer();
    })
    .then(function(buf){
      var wb=XLSX.read(new Uint8Array(buf), {type:'array'});
      var prod=sheetRows(wb, DATA_TAB);
      if(!prod) throw new Error('в файле нет листа «'+DATA_TAB+'»');
      readCats(sheetRows(wb, 'Категории'));
      readGlossary(sheetRows(wb, 'Справочник UZ'));
      var list=rowsToProducts(prod);
      if(!list.length) throw new Error('на листе «'+DATA_TAB+'» не найдено ни одного товара');
      P=list; prepare(); READY=true; markSource(); route();
    })
    .catch(function(e){ fail((e && e.message) || 'файл недоступен'); });
}

function markSource(){
  var el=document.getElementById('f-src');
  if(el) el.textContent=(LANG==='uz'?'Maʼlumotlar '+DATA_FILE+' dan · ':'Данные из '+DATA_FILE+' · ')
    + P.length + (LANG==='uz'?' mahsulot':' позиций');
}

function showLoader(){
  main.innerHTML='<div class="shell"><div class="empty" style="padding:120px 20px">'
    + '<i class="ti ti-loader-2 spin" aria-hidden="true"></i>'
    + '<b>'+(LANG==='uz'?'Katalog yuklanmoqda':'Загружаем каталог')+'</b></div></div>';
}

function fail(msg){
  READY=false;
  console.error('Каталог:', msg);
  main.innerHTML='<div class="shell"><div class="empty" style="padding:100px 20px">'
    + '<i class="ti ti-database-off" aria-hidden="true"></i>'
    + '<b>'+(LANG==='uz'?'Maʼlumotlarni yuklab boʻlmadi':'Не удалось загрузить данные')+'</b>'
    + '<span>'+esc(msg)+'</span>'
    + '<p style="margin-top:22px"><button class="btn-outline" id="retry">'
    + '<i class="ti ti-refresh" aria-hidden="true"></i> '+(LANG==='uz'?'Qayta urinish':'Повторить')+'</button></p>'
    + '<p style="margin-top:14px;font-size:13px;color:var(--faint)">'
    + (LANG==='uz'
        ? DATA_FILE+' fayli index.html yonida turishi kerak'
        : 'Файл '+DATA_FILE+' должен лежать рядом с index.html')
    + '</p></div></div>';
  var b=document.getElementById('retry');
  if(b) b.addEventListener('click', function(){ showLoader(); loadData(); });
}

/* ---------------- global ---------------- */
document.addEventListener('keydown',function(e){
  if(e.key==='Escape'){ closeQR(); return; }
  if(e.key==='/' && !/^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement.tagName)){
    if(window._slash){ e.preventDefault(); window._slash(); }
  }
});
var prog=document.getElementById('prog'), totop=document.getElementById('totop');
totop.addEventListener('click',function(){ window.scrollTo({top:0,behavior:'smooth'}); });
window.addEventListener('scroll',function(){
  var y=window.scrollY, max=document.body.scrollHeight-window.innerHeight;
  prog.style.width=(max>0?(y/max*100):0)+'%';
  totop.classList.toggle('show', y>700);
  if(topbar.classList.contains('on-dark')) topbar.classList.toggle('stuck', y>240);
  if(current) pbar.classList.toggle('show', y>430);
  if(window._parallax) window._parallax(y);
},{passive:true});

function route(){
  var h=location.hash.replace(/^#/,'')||'/';
  var qi=h.indexOf('?');
  var path=qi>-1?h.slice(0,qi):h;
  var params=new URLSearchParams(qi>-1?h.slice(qi+1):'');
  var lp=params.get('lang');
  if(lp==='uz'||lp==='ru') LANG=lp;
  closeQR();
  current=null; pbar.classList.remove('show');
  window._slash=null; window._parallax=null;
  paintChrome();
  document.title=(LANG==='uz'?'mahsulot katalogi':'каталог продукции');
  topbar.classList.remove('stuck');
  if(path.indexOf('/p/')===0) viewProduct(decodeURIComponent(path.slice(3)));
  else if(path==='/catalog') viewCatalog(params);
  else if(path==='/qr') viewSheet();
  else viewHome();
  if(path==='/catalog' && lastCatalogState && lastCatalogState.hash===location.hash){
    window.scrollTo(0, lastCatalogState.scrollY);
  } else {
    window.scrollTo(0,0);
  }
}
window.addEventListener('hashchange',function(e){
  if(!READY) return;
  try{
    var oldHash=new URL(e.oldURL).hash;
    if(oldHash.indexOf('#/catalog')===0){
      lastCatalogState={hash:oldHash, scrollY:window.scrollY};
    }
  }catch(err){}
  route();
});
paintChrome();
showLoader();
loadData();