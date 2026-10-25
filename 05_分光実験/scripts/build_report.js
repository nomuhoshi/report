const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Tab, TabStopType, AlignmentType, LineRuleType,
  Math: DMath, MathRun, MathFraction, MathSubScript, MathSuperScript, MathFunction,
  MathRoundBrackets, MathCurlyBrackets,
  Table, TableRow, TableCell, WidthType, ShadingType, ImageRun,
} = require('docx');

const FONT = { ascii: '游明朝', eastAsia: '游明朝', hAnsi: '游明朝', cs: '游明朝' };
const SZ = 21;                 // 10.5 pt
const TEXT_LINE = 304;         // 10.5 pt x 1.45 ≒ 15.2 pt（固定値）
const RED = 'FF0000';

// ---------- 文字列 → TextRun（x_{0} は下付き、^{2} は上付き）
function runs(str, opt = {}) {
  const out = [];
  const re = /(_\{[^}]*\}|\^\{[^}]*\})/g;
  str.split(re).forEach((part) => {
    if (!part) return;
    const base = { font: FONT, size: opt.size || SZ, bold: !!opt.bold, color: opt.color, italics: false };
    if (part.startsWith('_{')) out.push(new TextRun({ ...base, text: part.slice(2, -1), subScript: true }));
    else if (part.startsWith('^{')) out.push(new TextRun({ ...base, text: part.slice(2, -1), superScript: true }));
    else out.push(new TextRun({ ...base, text: part }));
  });
  return out;
}

// ---------- 段落
const para = (str, o = {}) => new Paragraph({
  children: runs(str, o),
  alignment: o.align,
  spacing: { before: o.before || 0, after: o.after || 0, line: o.line || TEXT_LINE, lineRule: LineRuleType.EXACT },
  indent: o.indent,
  
});
// 章標題 14 pt 太字、小標題 12 pt 太字（見出しスタイルは使わない）
const chapter = (t) => para(t, { bold: true, size: 28, line: 406, before: 200 });
const section = (t) => para(t, { bold: true, size: 24, line: 348, before: 80 });
const body = (t) => para(t);
const bullet = (t) => para('・' + t, { indent: { left: 210, hanging: 210 } });
const step = (t) => para(t, { indent: { left: 315, hanging: 315 } });
const red = (t) => para(t, { color: RED, indent: { left: 0 } });

// ---------- 数式（単倍行間、中央に式・右端に式番号）
const r = (t) => new MathRun(t);
const sub = (b, s) => new MathSubScript({ children: [r(b)], subScript: [r(s)] });
const sup = (b, s) => new MathSuperScript({ children: [r(b)], superScript: [r(s)] });
const fn = (name, args) => new MathFunction({ name: [r(name)], children: args });
const frac = (n, d) => new MathFraction({ numerator: n, denominator: d });
const paren = (c) => new MathRoundBrackets({ children: c });
const curly = (c) => new MathCurlyBrackets({ children: c });
let eqNo = 0;
function eq(children) {
  eqNo += 1;
  return new Paragraph({
    tabStops: [{ type: TabStopType.CENTER, position: 4536 }, { type: TabStopType.RIGHT, position: 9072 }],
    spacing: { before: 60, after: 60, line: 240, lineRule: LineRuleType.AUTO },
    children: [
      new TextRun({ children: [new Tab()], font: FONT, size: SZ }),
      new DMath({ children }),
      new TextRun({ children: [new Tab(), `（${eqNo}）`], font: FONT, size: SZ }),
    ],
  });
}

// ---------- 図の挿入位置（手描きの図を後で貼る）
const figNo = { n: 0 };
function figure(caption) {
  figNo.n += 1;
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 0, line: 240, lineRule: LineRuleType.AUTO },
      children: runs('★ここに手描きの図を挿入（この行は提出前に削除する）★', { color: RED, bold: true }),
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80, line: 240, lineRule: LineRuleType.AUTO },
      children: runs(`図${figNo.n}　${caption}`),
    }),
  ];
}


// ---------- 表（表名は表の上）
const CW = 9072;                       // A4 - 左右余白
let tblNo = 0;
function tcell(text, w, o = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: undefined,
    margins: { top: 30, bottom: 30, left: 50, right: 50 },
    children: [new Paragraph({
      alignment: o.left ? AlignmentType.LEFT : AlignmentType.CENTER,
      spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO },
      children: runs(String(text), { bold: !!o.head }),
    })],
  });
}
function table(caption, headers, rows, widths) {
  tblNo += 1;
  const out = [];
  out.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 40, line: 240, lineRule: LineRuleType.AUTO },
    children: runs('表' + tblNo + '　' + caption),
  }));
  out.push(new Table({
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((h, i) => tcell(h, widths[i], { head: true })) }),
      ...rows.map((r) => new TableRow({ cantSplit: true, children: r.map((v, i) => tcell(v, widths[i])) })),
    ],
  }));
  out.push(new Paragraph({ spacing: { before: 0, after: 60, line: 120, lineRule: LineRuleType.EXACT }, children: [] }));
  return out;
}
// ---------- 画像を貼る図（図名は図の下）
const PX = (cm) => Math.round(cm / 2.54 * 96);
function imgfig(file, caption, cm, cmH) {
  figNo.n += 1;
  const data = fs.readFileSync(__dirname + '/img/' + file);
  const dim = require('child_process');
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 0, line: 240, lineRule: LineRuleType.AUTO },
      children: [new ImageRun({ type: file.endsWith('.png') ? 'png' : 'jpg', data, transformation: { width: PX(cm), height: PX(cmH !== undefined ? cmH : cm * ASPECT[file]) } })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120, line: 240, lineRule: LineRuleType.AUTO },
      children: runs('図' + figNo.n + '　' + caption),
    }),
  ];
}
const ASPECT = { 'fig1_new.png': 1273 / 1400, 'fig2_new.png': 1512 / 1400, 'fig3.jpg': 1080 / 1592 };

// ---------- Excel から貼る図（図名は図の下）
function exfig(caption, where) {
  figNo.n += 1;
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 0, line: 240, lineRule: LineRuleType.AUTO },
      children: runs('★ここに Excel「' + where + '」のグラフを貼る（この行は提出前に削除する）★', { color: RED, bold: true }),
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 100, line: 240, lineRule: LineRuleType.AUTO },
      children: runs('図' + figNo.n + '　' + caption),
    }),
  ];
}

// ================= 本文 =================
const c = [];
c.push(para('5．分光実験', { bold: true, size: 28, line: 406, align: AlignmentType.CENTER, after: 60 }));

// 1 目的
c.push(chapter('1.目的'));
c.push(bullet('回折格子を用いて光の回折現象を理解し、Na、Hg、Cd ランプの輝線の波長を求める。'));
c.push(bullet('最小偏角を測定し、プリズム（ガラス）の屈折率分散を調べる。'));
c.push(bullet('反射型回折格子を用いたスペクトル測定の原理を理解し、Cd ランプと白色 LED のスペクトルを測定する。'));

// 2 原理
c.push(chapter('2.原理'));
c.push(section('2.1 回折格子による波長の測定'));
c.push(body('分光とは、光を波長ごとに分けることである。反射型回折格子は、溝が等間隔に並んだ鏡のようなものである。隣り合う溝で反射した光が干渉して、ある決まった方向だけ明るくなる。'));
c.push(body('溝の間隔を d、入射角を θ_{i}、回折角を θ_{m} とする（図1）。隣り合う 2 本の光の道のりの差は d(sinθ_{m} + sinθ_{i}) になり、これが波長 λ の整数倍のときに 2 本の光が重なって明るくなる。'));
c.push(eq([
  r('d'), paren([fn('sin', [sub('θ', 'm')]), r('+'), fn('sin', [sub('θ', 'i')])]),
  r('='), r('mλ'), r('   '), paren([r('m=0,±1,±2,…')]),
]));
c.push(body('ここで m は回折の次数である。本実験では光を垂直に入射し（θ_{i} = 0）、1 次光（m = +1）だけを観測するので、式（1）は次式になる。'));
c.push(eq([r('d'), fn('sin', [r('θ')]), r('='), r('λ')]));
c.push(body('使う回折格子は 1200 lines/mm なので、d = 1/1200 mm である。スクリーンの目盛の原点は光軸から x_{0} の位置にあり、輝線の読みを x_{1}、回折格子からスクリーンまでの距離を L とすると、回折角 θ は次の式で求まる。'));
c.push(eq([fn('tan', [r('θ')]), r('='), frac([sub('x', '0'), r('+'), sub('x', '1')], [r('L')])]));
c.push(body('式（3）から θ を求め、式（2）に代入すると、輝線の波長 λ が得られる。'));
c.push(...imgfig('fig1_new.png', '反射型回折格子での回折角測定', 11.0));

c.push(section('2.2 プリズムによる屈折率の測定'));
c.push(body('頂角 α のプリズムに入った光は、2 つの面で屈折して、入ってきた方向から偏角 δ だけ曲がる。入る側の入射角と屈折角を i、r、出る側を i′、r′ とすると（図2）、次の関係が成り立つ。'));
c.push(eq([r('δ'), r('='), paren([r('i−r')]), r('+'), paren([r('i′−r′')])]));
c.push(eq([r('α'), r('='), r('r+r′')]));
c.push(eq([r('n'), r('='), frac([fn('sin', [r('i')])], [fn('sin', [r('r')])]), r('='), frac([fn('sin', [r('i′')])], [fn('sin', [r('r′')])])]));
c.push(body('n は屈折率であり、光の波長によって値が異なる（屈折率分散）。δ は r = r′、i = i′ のとき最小値 δ_{0} をとる（最小偏角）。このとき式（4）、（5）より r = α/2、i = (δ_{0} + α)/2 となるので、式（6）から次式が得られる。'));
c.push(eq([
  r('n'), r('='),
  frac([fn('sin', [frac([sub('δ', '0'), r('+'), r('α')], [r('2')])])], [fn('sin', [frac([r('α')], [r('2')])])]),
]));
c.push(body('本実験ではプリズムは正三角形で α = 60° である。δ_{0} は、プリズムからスクリーンまでの距離を L、屈折光がスクリーンに届く点のプリズムからの水平距離を x として、次式から求める。'));
c.push(eq([fn('tan', [sub('δ', '0')]), r('='), frac([r('L')], [r('x')])]));
c.push(...imgfig('fig2_new.png', 'プリズムによる偏角の測定', 9.0));

c.push(section('2.3 反射型回折格子を用いたスペクトル測定'));
c.push(body('入ってくる光の道すじと、検出器へ向かう回折光の道すじがつくる角 φ を固定して、回折格子を回転させる。回折格子を入射角 0° の位置から θ_{r} だけ回転させたとき、検出される回折光の条件は次の式になる。'));
c.push(eq([
  r('d'), curly([fn('sin', [paren([sub('θ', 'r'), r('+'), r('φ')])]), r('+'), fn('sin', [sub('θ', 'r')])]),
  r('='), r('mλ'),
]));
c.push(body('和積の公式より sin(θ_{r} + φ) + sinθ_{r} = 2 sin(θ_{r} + φ/2) cos(φ/2) なので、式（9）は次のように変形できる。'));
c.push(eq([
  fn('sin', [paren([sub('θ', 'r'), r('+'), frac([r('φ')], [r('2')])])]), r('='),
  frac([r('mλ')], [r('2d'), fn('cos', [frac([r('φ')], [r('2')])])]),
]));
c.push(body('φ と d は一定なので、θ_{r} を変えると検出器に届く光の波長が変わる。回折格子を回転させながら光の強さを測れば、スペクトルが得られる。'));
c.push(body('本実験では回転量をマイクロメータの読み x [目盛] で測った。θ_{r} は x にほぼ比例すると考えられるので、θ_{r} = k x + α と書ける。これを式（10）に入れて整理すると、λ と x の関係は次の二次関数で表せる（A、B、C は係数）。k、α、φ の値は分からないが、波長が分かっている Cd ランプの輝線を使えば A、B、C は決められる。'));
c.push(eq([r('λ'), r('='), sup('Ax', '2'), r('+'), r('Bx'), r('+'), r('C')]));

// 3 実験の方法
c.push(chapter('3.実験の方法'));
c.push(section('3.1 実験器材'));
c.push(bullet('光源：Na ランプ、Hg ランプ、Cd ランプ、白色 LED'));
c.push(bullet('分光素子：反射型回折格子（Thorlabs GR25-1205、1200 lines/mm、ブレーズ波長 500 nm）、プリズム（正三角形ガラス、頂角 60°）'));
c.push(bullet('光学素子：ブレッドボード、波長可変スリット、両凸レンズ、平凸レンズ、ビームブロッカー、スクリーン、回転ステージ'));
c.push(bullet('光検出器：増幅フォトダイオード（Thorlabs DET36A、幅 200 µm のスリット付き）、デジタルマルチメータ'));
c.push(body('光学系を図3に示す。'));
c.push(...imgfig('fig3.jpg', '光学系（真上から見た図）', 15.5));

c.push(section('3.2 実験1：各種ランプの波長の決定'));
c.push(step('①　回転ステージに反射型回折格子を設置し、白色 LED の光を垂直に入射する（0次光が入射光と重なるようにする）。'));
c.push(step('②　1次回折光がスクリーンに映るようにスクリーンを固定し、回折格子からスクリーンまでの距離 L を測る（L = 150 mm）。'));
c.push(step('③　光軸からスクリーンの目盛の原点までの距離 x_{0} を測る（x_{0} = 50 mm）。'));
c.push(step('④　光源を Na、Hg、Cd ランプの順に替え、各輝線の位置 x_{1}（左端、中心、右端）を読み取り、輝線の色も記録する。'));
c.push(step('⑤　Cd ランプの赤線は左右の端が読み取りにくく中心のみ、青線は光が細く左端のみ読み取った。'));

c.push(section('3.3 実験2：最小偏角 δ_{0} の測定'));
c.push(step('①　プリズムを置き、白色 LED の光を当てる。可変スリットの幅は 200 µm にする。'));
c.push(step('②　プリズムを回転させて偏角が最小になるようにし、スクリーンを固定して、プリズムからスクリーンまでの距離 L を測る。'));
c.push(step('③　光源を Cd ランプに替える。各輝線（赤、緑、水色、青）について、x が最大（偏角が最小）になるようにプリズムの角度を微調整する。'));
c.push(step('④　各輝線の x を測る。'));
c.push(step('⑤　スクリーンの位置を変え（L = 150、175、200 mm）、計 3 回測定する。'));

c.push(section('3.4 実験3：反射型回折格子を用いたスペクトル測定'));
c.push(step('①　回転ステージに反射型回折格子を固定し、白色光を垂直に入射する。Ⓐの位置に平凸レンズ（回折格子側を凸）、Ⓑの位置にスリット付きフォトダイオードを置く。'));
c.push(step('②　白色 LED を点灯し、紙をかざして回折光の色を確かめながら、赤色の光がスリットに入る位置に回転ステージを合わせた。この状態で回転ステージ以外を固定した。'));
c.push(step('③　回転ステージの回転量はマイクロメータで読み取る。1 目盛が 0.01 mm、1 回転が 0.5 mm である。以下、マイクロメータの読みを x [目盛] と書く。'));
c.push(step('④　先に白色 LED のスペクトルを測定した（実験3-4）。x を 25 目盛ずつ、出力電圧が速く変わる範囲では 2 目盛ずつ変えながら、出力電圧を記録した。'));
c.push(step('⑤　光源を Cd ランプに変え、入射スリット幅を 0.1 mm として x = 220 から 743 まで掃引し、出力電圧を記録した（実験3-2）。輝線の付近では 1 目盛ずつ変えた。あわせて、スリットの位置で目視した光の色を記録した。'));
c.push(step('⑥　最も強い輝線について、x = 600 に戻して入射スリット幅を 1 mm に広げ、x = 605〜660 を 5 目盛ずつ掃引した（実験3-3）。'));
c.push(body('入射スリット幅は 0.1 mm と 1 mm の 2 通りとした。また、実験3-2 は赤線のみ 2 回、他の輝線は 1 回の掃引である。'));

// 4 実験の結果
c.push(chapter('4.実験の結果'));

c.push(section('4.1 各ランプの輝線の波長'));
c.push(body('反射型回折格子による 1 次回折光のスクリーン上の位置から、各ランプの輝線の波長を求めた。x0 = 50 mm、L = 150 mm である。結果を表1に示す。'));
c.push(...table('各ランプの輝線の波長',
 ['ランプ','色','x1 左 [mm]','x1 中 [mm]','x1 右 [mm]','x [mm]','θ [deg]','λ [nm]','文献値 [nm]','相対誤差 [%]'],
 [['Na','オレンジ','95.0','97.0','99.0','147.0','44.42','583.3','589.3','1.0'],
  ['Hg','オレンジ','87.5','89.0','90.9','139.0','42.82','566.4','579.1','2.2'],
  ['Hg','緑','74.2','75.5','77.0','125.5','39.92','534.7','546.1','2.1'],
  ['Hg','青','37.1','38.2','39.2','88.2','30.46','422.4','435.8','3.1'],
  ['Cd','赤','—','12.5','—','62.5','22.62','320.5','643.8','50.2'],
  ['Cd','緑','52.0','54.5','57.0','104.5','34.86','476.4','508.6','6.3'],
  ['Cd','水色','35.5','37.0','40.5','87.0','30.11','418.1','480.0','12.9'],
  ['Cd','青','35.6','—','—','85.6','29.71','413.0','467.8','11.7']],
 [760,1000,870,870,870,830,830,830,1100,1112]));
c.push(body('x1 は輝線の左端、中心、右端の読みであり、波長の計算には中心の値を使った。x は x0 + x1（中心）である。Cd の赤線は左右の端が読み取れなかったので中心だけ、Cd の青線は光が細かったので左端だけを読み、その値で計算した。読み取れなかったところは「—」とした。文献値は講義 付録の値で、Na は 588.995 nm と 589.592 nm の平均を使った。'));
c.push(body('Na は 1 本、Hg は 3 本、Cd は 4 本の輝線が見えた。Na と Hg の波長は文献値より 1.0〜3.1 % 小さく、Cd は 6.3〜50.2 % 小さかった。Cd の赤線は文献値の約半分になった。'));

c.push(section('4.2 プリズムの屈折率分散'));
c.push(body('Cd ランプの 4 本の輝線について、スクリーンの位置を 3 回変えて最小偏角 δ_{0} を測り、屈折率 n を求めた。波長は表1の文献値を使った。結果を表2に、3 回の平均値を表3に示す。波長と屈折率の関係を図4に示す。'));
c.push(...table('各スクリーン位置における最小偏角と屈折率', ['色','λ [nm]','L [mm]','x [mm]','δ_{0} [deg]','n'],
 [['赤','643.8','150','136','47.80','1.6160'],
  ['赤','643.8','175','155','48.47','1.6228'],
  ['赤','643.8','200','188','46.77','1.6053'],
  ['緑','508.6','150','128','49.52','1.6335'],
  ['緑','508.6','175','151','49.21','1.6304'],
  ['緑','508.6','200','181','47.85','1.6165'],
  ['水色','480.0','150','127','49.75','1.6358'],
  ['水色','480.0','175','149','49.59','1.6342'],
  ['水色','480.0','200','176','48.65','1.6247'],
  ['青','467.8','150','118','51.81','1.6562'],
  ['青','467.8','175','143','50.75','1.6457'],
  ['青','467.8','200','170','49.64','1.6346']],
 [1300,1400,1300,1400,1700,1972]));
c.push(...table('屈折率の平均値', ['色','λ [nm]','n（3 回の平均）'],
 [['赤','643.8','1.6147'],['緑','508.6','1.6268'],['水色','480.0','1.6315'],['青','467.8','1.6455']],
 [2500,2500,4072]));
c.push(...imgfig('fig4.png', '波長と屈折率の関係', 13.99, 8.99));
c.push(body('屈折率は 1.605〜1.656 であった。いずれのスクリーン位置でも、波長が短いほど屈折率は大きかった。'));

c.push(section('4.3 Cd ランプのスペクトルと波長校正曲線'));
c.push(body('入射スリット幅を 0.1 mm にして、回転ステージを回しながら Cd ランプの 1 次回折光の出力電圧を測った。結果を図5に示す。'));
c.push(...imgfig('fig5.png', 'Cd ランプのスペクトル（入射スリット幅 0.1 mm）', 14.01, 8.48));
c.push(body('x = 220〜743 の範囲で 4 つのピークが見えた。それぞれのピークの位置と、スリットのところで目で見た色から決めた文献波長を表4に示す。'));
c.push(...table('Cd ランプの輝線のピーク位置', ['色','ピーク位置 x [目盛]','ピーク出力 [V]','文献波長 [nm]'],
 [['赤','278.5','0.0070','643.847'],
  ['緑','634','0.0089','508.582'],
  ['水色','705','0.0075','479.992'],
  ['青','738','0.0059','466.235 / 467.815']],
 [1500,2500,2100,2972]));
c.push(body('赤線は x = 278 と 279 の出力電圧が同じだったので、その中点をピーク位置とした。'));
c.push(body('赤・緑・水色の 3 本を用いて λ = A x² + B x + C の係数を最小二乗法で求め、A = −5.201×10⁻⁵ nm/目盛²、B = −0.3330 nm/目盛、C = 740.63 nm を得た。青線は 2 本の輝線が重なっているため、校正には使わなかった。校正曲線を図6に示す。'));
c.push(...imgfig('fig6.png', '波長校正曲線', 14.01, 8.48));
c.push(body('この校正曲線で x = 738 のピークを計算すると λ = 466.5 nm になった。文献値 467.815 nm と 466.235 nm に対応する位置はそれぞれ x = 734.9 と 738.7 なので、見えたピークはその間にあった。この 2 本は 1 つのピークにしか見えず、分けることができなかった。'));

c.push(section('4.4 スリット幅とスペクトル幅'));
c.push(body('4.3 で一番強かった緑線について、入射スリット幅を 0.1 mm と 1 mm に変えて同じ輝線を掃引した。スペクトル幅は、ピークの高さの半分のところで測った幅（半値幅）とした。波長への換算には 4.3 の校正曲線を使った。結果を表5に、スリット幅とスペクトル幅の関係を図7に示す。'));
c.push(...table('入射スリット幅とスペクトル幅（Cd 緑線 508.582 nm）', ['入射スリット幅 [mm]','ピーク出力 [V]','半値幅 [nm]'],
 [['0.1','0.0089','4.4'],['1','0.0162','11.6']],
 [3024,3024,3024]));
c.push(...imgfig('fig7.png', '入射スリット幅とスペクトル幅の関係', 13.99, 8.50));
c.push(body('スリット幅を 0.1 mm から 1 mm にすると、ピーク出力は 0.0089 V から 0.0162 V に増え、半値幅は 4.4 nm から 11.6 nm に広がった。'));
c.push(body('なお、スリット幅 1 mm の測定は x = 605〜660 の範囲だけで、ピークから離れたところの値は測っていない。背景には、測った範囲の両端の平均値を使った。'));

c.push(section('4.5 白色 LED のスペクトル'));
c.push(body('4.3 の校正曲線を使って、白色 LED の出力電圧を波長に対して並べ直した。結果を図8に、ピーク波長と半値幅を表6に示す。'));
c.push(...imgfig('fig8.png', '白色 LED のスペクトル', 16.00, 9.50));
c.push(...table('白色 LED のピーク波長とスペクトル幅', ['山','ピーク波長 [nm]','ピーク出力 [V]','半値幅 [nm]'],
 [['長波長側','580.2','0.1988','134'],['短波長側','450.0','0.1352','18']],
 [2772,2100,2100,2100]));
c.push(body('測れた範囲は 424〜672 nm であった。グラフには山が 2 つあり、その間の谷は 482〜492 nm のあたりであった。半値幅は、長波長側の山は背景を 0 として、短波長側の山は谷の 0.0312 V を背景として求めた。'));

// 5 課題
c.push(chapter('5.課題'));

c.push(section('5.1 光の波長と色の関係'));
c.push(body('本実験で見えた輝線の色と波長を、波長の長い順に表7に示す。波長は表1の文献値である。'));
c.push(...table('観測した輝線の色と波長', ['色','λ [nm]','光源'],
 [['赤','643.8','Cd'],['オレンジ','589.3','Na'],['オレンジ','579.1','Hg'],['緑','546.1','Hg'],
  ['緑','508.6','Cd'],['水色','480.0','Cd'],['青','467.8','Cd'],['青','435.8','Hg']],
 [2000,2500,4572]));
c.push(body('見えた色は、波長が長い順に 赤、オレンジ、緑、水色、青 であった。つまり、波長が長いほど赤に近く、短いほど青に近い色に見える。'));
c.push(body('また実験1では、白色 LED の光は 1 本の輝線ではなく、スクリーンの上で赤から紫まで続けて変わって見えた。これは白色 LED が、決まった波長だけでなく広い範囲の波長を含んでいるからだと考えられる。'));

c.push(section('5.2 輝線と白色 LED のスペクトル幅'));
c.push(body('4.4 と 4.5 で求めたスペクトル幅を表8に示す。いずれも半値幅で定義した。'));
c.push(...table('輝線と白色 LED のスペクトル幅の比較', ['光源','ピーク波長 [nm]','半値幅 [nm]'],
 [['Cd 緑線','508.6','4.4'],
  ['白色 LED（長波長側）','580.2','134'],
  ['白色 LED（短波長側）','450.0','18']],
 [3000,3000,3072]));
c.push(body('Cd の緑線の半値幅は 4.4 nm であった。4.3 のとおり、間隔が 1.58 nm しかない 2 本の輝線（466.235 nm と 467.815 nm）が 1 つのピークにしか見えなかったので、この 4.4 nm は装置の分解能に近い値だと考えられる。つまり Cd の輝線の本当の幅は 4.4 nm よりもっと狭いと思われる。一方、白色 LED の長波長側の山は 134 nm で、4.4 nm よりかなり広い。'));
c.push(body('ランプの光は、原子の決まった 2 つのエネルギー準位の間を電子が移るときに出る。準位が決まっていて、その差のエネルギーも決まっているので、出てくる光の波長も決まった値になる。これが輝線の幅が狭い理由だと考えられる。'));
c.push(body('白色 LED の光は、半導体から出る光と、それを吸収した蛍光体から出る光が混ざったものである。どちらも 1 つの決まった準位からではなく、幅のあるエネルギーの範囲から出るので、スペクトル幅は広くなると考えられる。図8 に山が 2 つあったのは、この 2 つの光に対応している可能性が考えられる。'));
c.push(body('なお Na と Hg の輝線については、実験3 でスペクトル幅を測っていない。ただし実験1 では、Na・Hg・Cd はどれもスクリーンの上に線として見えたのに対し、白色 LED は続けて変わって見えた。'));

c.push(section('5.3 スペクトル分解能'));
c.push(body('表5 のとおり、入射スリットを 0.1 mm から 1 mm に広げると、半値幅は 4.4 nm から 11.6 nm になった。スリットを広げると分解能は悪くなる。'));
c.push(body('実際の例として、4.3 の Cd の青線がある。466.235 nm と 467.815 nm の間隔は 1.58 nm で、半値幅 4.4 nm より狭い。そのため 2 本は分かれず、x = 738 に 1 つのピークとして見えた。2 本を分けるには、半値幅がその間隔より狭くならないといけないと考えられる。'));
c.push(body('分解能を良くする方法として、次の 2 つが考えられる。'));
c.push(step('①　入射スリットを狭くする。表5 でスリットを 1/10 にすると半値幅が約 1/2.6 になったので、もっと狭くすれば半値幅はもっと小さくなると考えられる。ただしピーク出力も一緒に下がる（0.0162 V → 0.0089 V）ので、信号が小さくなりすぎない範囲にする必要がある。'));
c.push(step('②　回折次数 m を上げる。式(5) の sin θ_{m} = mλ/d を見ると、波長が少し違う 2 つの光の回折角の差は、m が大きいほど大きくなる。つまり次数を上げれば 2 本を分けやすくなると考えられる。'));
c.push(body('なお、本実験では入射スリット幅を 0.1 mm と 1 mm の 2 つでしか測っていないので、もっと狭くしたときにどこまで半値幅が小さくなるかは確かめていない。'));

// 6 考察
c.push(chapter('6.考察'));
c.push(para('★この章は自分で書く。書いたらこの行を削除する★', { color: RED, bold: true }));

// 7 結論
c.push(chapter('7.結論'));
c.push(bullet('反射型回折格子の 1 次回折光から、各ランプの輝線の波長を求めた。Na と Hg は文献値と 1.0〜3.1 % の差だったが、Cd は 6.3〜50.2 % の差があった。'));
c.push(bullet('最小偏角から求めたガラスの屈折率は 1.605〜1.656 であり、波長が短いほど屈折率が大きいという屈折率分散を確認した。'));
c.push(bullet('Cd の輝線で波長を校正した装置（半値幅 4.4 nm）で白色 LED のスペクトルを測り、450.0 nm と 580.2 nm に山があることが分かった。'));

// 8 参考文献
c.push(chapter('8.参考文献'));
c.push(para('［1］　東京理科大学先進工学部物理工学科　物理学実験（2004 年 3 月改訂）「5．分光実験」。'));
c.push(para('［2］　理科年表（輝線の波長は、［1］の付録に記載された値を用いた）。'));

const doc = new Document({
  creator: '',
  title: '5．分光実験',
  styles: {
    default: {
      document: {
        run: { font: FONT, size: SZ },
        paragraph: { spacing: { before: 0, after: 0, line: TEXT_LINE, lineRule: LineRuleType.EXACT } },
      },
    },
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1417, bottom: 1417, left: 1417, right: 1417 } } },
    children: c,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = process.argv[2];
  fs.writeFileSync(out, buf);
  console.log('wrote', out);
});
