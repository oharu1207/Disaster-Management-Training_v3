/* ============================================================
   ICS学習支援システム app.js v3.0
   PowerPoint風ドラッグ矢印描画
   ============================================================ */
(() => {

  // ================================================================
  // SCENARIO & DATA
  // ================================================================
  const SCENARIO = {
    id: "s1-1",
    title: "急性期：医療救護活動の組織構造を可視化せよ",
    phase: "急性期"
  };

  // ================================================================
  // PHASE 定数
  // state.phase の値がフェーズ名で分かるようにする。
  // DOM 上の .phase-view の並び順やインデックスには依存しない。
  // ================================================================
  const PHASE = {
    ORIENTATION:      0,   // オリエンテーション
    ACUTE_MAP:        1,   // マップ構築（急性期）
    ACUTE_COMPARE:    2,   // 比較・分析（急性期）
    ACUTE_RECORD:     3,   // 対応検証記録レビュー（急性期）
    RECOVERY_PREP:    4,   // 復旧期準備
    RECOVERY_MAP:     5,   // 復旧期マップ
    RECOVERY_COMPARE: 6,   // 比較・分析（復旧期）
    RECOVERY_RECORD:  7,   // 対応検証記録（復旧期）
    SEQUENCE:         8,   // シーケンス図構築
  };

  const STORAGE_KEY            = "ics-learning-system-v1";
  const STORAGE_SCHEMA_VERSION = 1;

  // フェーズ番号 → 対応する .phase-view の HTML id
  // DOM 順ではなく id で直接参照するため、並び替えに強い。
  // [CHANGED] キー（数値）を PHASE 定数に合わせて更新。値（DOM id 文字列）は変更しない。
  const PHASE_VIEW_ID = {
    [PHASE.ORIENTATION]:      "phase-0",
    [PHASE.ACUTE_MAP]:        "phase-1",
    [PHASE.ACUTE_COMPARE]:    "phase-2",
    [PHASE.ACUTE_RECORD]:     "phase-acuteRecord",
    [PHASE.RECOVERY_PREP]:    "phase-5",
    [PHASE.RECOVERY_MAP]:     "phase-6",
    [PHASE.RECOVERY_COMPARE]: "phase-recoveryCompare",
    [PHASE.RECOVERY_RECORD]:  "phase-recoveryRecord",
    [PHASE.SEQUENCE]:         "phase-3",
  };

  // === ノード一覧（Excelノート_.xlsx より） ===
  const PALETTE_NODES = [
    { label: "県庁",                      group: "g-command", icon: "🏛" },
    { label: "C県A保健所",               group: "g-command", icon: "🏥" },
    { label: "地域災害医療コーディネーター", group: "g-command", icon: "🩺" },
    { label: "DMAT",                     group: "g-unit",    icon: "🚑" },  // [CHANGED] g-section→g-unit
    { label: "DHEAT",                    group: "g-section", icon: "📋" },
    { label: "DPAT",                     group: "g-unit",    icon: "🧠" },  // [CHANGED] g-section→g-unit
    { label: "DWAT",                     group: "g-unit",    icon: "💧" },  // [CHANGED] g-section→g-unit
    { label: "JMAT",                     group: "g-unit",    icon: "🏥" },  // [CHANGED] g-section→g-unit
    { label: "JRAT",                     group: "g-unit",    icon: "🔧" },  // [CHANGED] g-section→g-unit
    { label: "C県看護協会",              group: "g-unit",    icon: "💉" },  // [CHANGED] g-section→g-unit
    { label: "医師会",                   group: "g-unit",    icon: "👨‍⚕️" },
    { label: "歯科医師会",               group: "g-unit",    icon: "🦷" },
    { label: "AB薬剤師会",               group: "g-unit",    icon: "💊" },
    { label: "C県栄養士会",              group: "g-unit",    icon: "🥗" },
    { label: "地域包括支援センター",      group: "g-unit",    icon: "🤝" },
    { label: "W民間団体",                group: "g-unit",    icon: "🏢" },
    { label: "市町村保健センター",        group: "g-unit",    icon: "🏘" },  // [CHANGED] g-team→g-unit
    { label: "避難所",                   group: "g-team",    icon: "🏘" },
    { label: "医療機関",                 group: "g-team",    icon: "🏥" },
  ];

  // === エッジ種別（3種類のみ） ===
  // label: エッジラベル, stroke: 線色, bidirectional: 双方向フラグ
  const EDGE_TYPES = [
    { label: "指示命令", desc: "命令・指示の方向を設定", stroke: "#ff6b6b", bidirectional: false },
    { label: "情報伝達", desc: "情報の流れる方向を設定", stroke: "#4d8fff", bidirectional: false },
    { label: "連携協力", desc: "双方向・協力関係（固定）", stroke: "#3dcf8a", bidirectional: true  },
  ];

  // label → {stroke, bidirectional} の逆引きマップ
  const EDGE_MAP = Object.fromEntries(
    [...EDGE_TYPES, { label: "支援", stroke: "#c084fc", bidirectional: false }]
      .map(t => [t.label, t])
  );


  // === 復旧期データ ===
  const RECOVERY_PALETTE_NODES = [
    { label: "県庁",                          group: "g-command", icon: "🏛" },
    { label: "C県A保健所",                   group: "g-command", icon: "🏥" },
    { label: "地域災害医療コーディネーター", group: "g-command", icon: "🩺" },
    { label: "DHEAT",                         group: "g-section", icon: "📋" },  // g-section 維持（撤退候補として学習者が判断）
    { label: "DPAT",                          group: "g-unit",    icon: "🧠" },  // [CHANGED] g-section→g-unit
    { label: "DWAT",                          group: "g-unit",    icon: "💧" },  // [CHANGED] g-section→g-unit
    { label: "JMAT",                          group: "g-unit",    icon: "🏥" },  // [CHANGED] g-section→g-unit
    { label: "JRAT",                          group: "g-unit",    icon: "🔧" },  // [CHANGED] g-section→g-unit
    { label: "C県看護協会",                   group: "g-unit",    icon: "💉" },  // [CHANGED] g-section→g-unit
    { label: "DCAT",                         group: "g-unit",    icon: "🤝" },  // [CHANGED] g-section→g-unit
    { label: "医師会",                        group: "g-unit",    icon: "👨‍⚕️" },
    { label: "歯科医師会",                    group: "g-unit",    icon: "🦷" },
    { label: "AB薬剤師会",                    group: "g-unit",    icon: "💊" },
    { label: "C県栄養士会",                   group: "g-unit",    icon: "🥗" },
    { label: "地域包括支援センター",          group: "g-unit",    icon: "🤝" },
    { label: "W民間団体",                     group: "g-unit",    icon: "🏢" },
    { label: "社会福祉士会",                 group: "g-unit",    icon: "👥" },
    { label: "地域支え合いセンター",         group: "g-unit",    icon: "🏘" },
    { label: "介護支援専門員協会",            group: "g-unit",    icon: "🧑‍⚕️" },
    { label: "市町村保健センター",            group: "g-unit",    icon: "🏘" },  // [CHANGED] g-team→g-unit
    { label: "避難所",                        group: "g-team",    icon: "🏘" },
    { label: "福祉避難所",                    group: "g-team",    icon: "🏡" },
    { label: "在宅避難者",                    group: "g-team",    icon: "🏠" },
    { label: "仮設住宅",                     group: "g-team",    icon: "🏠" },
  ];

  const RECOVERY_BENEFICIARY_LABELS = new Set(["避難所", "福祉避難所", "在宅避難者", "仮設住宅"]);

  const PHASE6_BENEFICIARY_LABELS = new Set(["避難所", "医療機関", "福祉避難所", "在宅避難者", "仮設住宅"]);

  const NODE_DESCRIPTIONS = {
    "県庁":                          "広域指揮・県全体の災害対応方針を決定する行政機関",
    "C県A保健所":                   "地域の現地指揮拠点・支援チームの受入・調整を担う行政機関",
    "地域災害医療コーディネーター": "医療資源の配分と機関間調整を専門的に支援する",
    "DMAT":                          "災害や新興感染症等のまん延時に，地域において必要な医療提供体制を支援し，傷病者の生命を守るため厚生労働省の認めた専門的な研修・訓練を受けた災害派遣医療チーム",
    "DHEAT":                         "一定規模以上の災害が発生した際に，被災都道府県庁の保健医療福祉調整本部及び保健所が担う指揮・総合調整機能等を支援するため、専門的な研修・訓練を受けた都道府県等の職員により構成される応援派遣チーム",
    "DPAT":                          "被災地域の専門性の高い精神科医療の提供と精神保健活動の支援（入院患者等の避難及び搬送，被災医療機関・災害ストレスへの支援等）を行うために，都道府県によって組織される，災害派遣精神医療チーム",
    "DWAT":                          "主に一般避難所における要配慮者等の二次被害の防止，安定的な日常生活への移行を支えることを目的に，多様な福祉職で構成する災害派遣福祉チーム",
    "JMAT":                          "被災者の生命及び健康を守り，被災地の公衆衛生を回復し，地域医療や地域包括ケアシステムの再生・復興を支援することを目的とする日本医師会災害医療チーム",
    "JRAT":                          "一般社団法人日本災害リハビリテーション支援協会．被災者・要配慮者の生活不活性発病や災害関連死等の予防に関する支援を行う",
    "C県看護協会":                   "都道府県看護協会に登録されている災害支援ナースの派遣・調整を担う職能団体",
    "医師会":                        "地域医療を担う医師の職能団体",
    "歯科医師会":                    "口腔ケア・身元確認等を担う歯科医師の職能団体",
    "AB薬剤師会":                    "薬剤管理・服薬支援を担う薬剤師の職能団体",
    "C県栄養士会":                   "避難所等での栄養管理・食支援を担う職能団体",
    "地域包括支援センター":           "高齢者・要支援者の生活支援ニーズを把握する機関",
    "W民間団体":                     "行政を補完するボランティア・NPO等の民間支援組織",
    "市町村保健センター":             "住民に最も近い保健活動の実施主体となる行政機関",
    "避難所":                        "自宅に居住できなくなった被災者を一時的に受け入れ保護するための場所",
    "医療機関":                      "急性期における傷病者受入・医療救護の実施場所",
    "DCAT":                          "災害発生時に要配慮者を支援するため，介護福祉士等による専門職で構成するチーム",
    "社会福祉士会":                  "社会福祉士の職能団体．日常生活の再建を支援するための相談援助と，諸 関係機関との連携・調整を行う",
    "地域支え合いセンター":           "仮設住宅や在宅等の被災者を巡回訪問し，困りごとやへの相談対応，交流の場づくりなどを支援する地域の拠点",
    "介護支援専門員協会":             "介護支援専門員の職能団体．大規模災害時に被災地へ災害支援ケアマネジャーを派遣し，高齢者の実態把握，避難所での支援活動を行う",
    "福祉避難所":                    "高齢者や障害者など特別な配慮を必要とする要配慮者を受け入れる避難所",
    "在宅避難者":                    "自宅で居住の継続ができる状況で，自宅に留まる被災住民",
    "仮設住宅":                      "災害で住まいを失った人に対し，行政が一時的に提供する無料の住宅",
  };

  const MAP_PHASE_CONFIG = {
    [PHASE.ACUTE_MAP]: {
      key: "acute", paletteNodes: PALETTE_NODES,
      beneficiaries: new Set(["避難所", "医療機関"]),
      domIds: { canvas:"canvas-acute", svg:"svgLayer-acute", palette:"palette-acute",
                wrap:"canvasWrap-acute", stat:"canvasStat-acute", hint:"arrowModeHint-acute" },
      markerSuffix: "-acute",
    },
    [PHASE.RECOVERY_PREP]: {
      key: "p5",
      isReadOnly: true,
      beneficiaries: new Set(["避難所", "医療機関"]),
      domIds: { canvas:"canvas-p5", svg:"svgLayer-p5",
                wrap:"canvasWrap-p5", stat:"canvasStat-p5" },
      markerSuffix: "-p5",
    },
    [PHASE.RECOVERY_MAP]: {
      key: "p6",
      paletteNodes: RECOVERY_PALETTE_NODES,
      beneficiaries: PHASE6_BENEFICIARY_LABELS,
      domIds: { canvas:"canvas-p6", svg:"svgLayer-p6", palette:"palette-p6",
                wrap:"canvasWrap-p6", stat:"canvasStat-p6", hint:"arrowModeHint-p6" },
      markerSuffix: "-p6",
    },
  };

  // ================================================================
  // ACUTE_RECORD_CONTENT — 対応検証記録フェーズのデータ定義 [ADDED]
  // 抜粋・問の増減・文言修正はここだけで完結する。
  // ================================================================
  const ACUTE_RECORD_CONTENT = {
    title: "対応検証記録（急性期 課題抜粋）",
    excerpts: [
      {
        id: "1",
        text: "県庁との情報共有については、県庁における窓口が統一されておらず、県庁の各課から同じような内容の確認が幾度となくあり、保健所が把握していない問題への対応依頼等があり、保健所は混乱することがあった。"
      },
      {
        id: "2",
        text: "県庁本部で、県庁と支援団体間だけで決められていた被災地支援活動などがあった。そういった活動の中には、保健所が現場ですでに取り組んでいた活動もあり、二重になってしまうこともあった。"
      },
      {
        id: "3",
        text: "「参加者が多いと、会議時間が長くなってしまう」「会議では情報共有はできたが、具体的な支援団体の配置や活動における役割分担などにまで話が及ぶことはあまりなかった」"
      },
      {
        id: "4",
        text: "DHEATを派遣するにあたり、保健所側の要望も聞いて欲しい。必ずしもプッシュ型である必要はあまりないのではないか。"
      },
      {
        id: "5",
        text: "「県庁保健医療調整本部‐保健所現地保健医療調整本部との連携は、かなり薄かったと言わざるを得なかった。保健所には県庁本部の動きはまったく伝わってこなかった。」「情報網が遮断されたこともあって、保健所の全体的な活動を本庁に伝える手段もなく、本庁から聞かれることもなかった。」"
      }
    ],
    questions: [
      {
        id: "q4",
        kind: "singleChoice",
        label: "問4．問3で指摘したICS原則違反は，対応検証記録の何番と対応するか1つ示せ。",
        optionsSource: "excerpts",
        required: true
      },
      {
        id: "q5",
        kind: "textarea",
        label: "問5．問2・問3での指摘を踏まえつつ、あなたの考える組織構造上の問題が対応失敗をどのように引き起こしたか説明せよ。",
        placeholder: "例）保健所が…という構造的問題があったため、…という失敗が生じた。",
        maxLength: 200,
        required: true
      }
    ]
  };

  // ================================================================
  // RECOVERY_COMPARE_CONTENT — 復旧期比較・分析フェーズのデータ定義 [ADDED]
  // ================================================================
  const RECOVERY_COMPARE_CONTENT = {
    questions: [
      {
        id: "q6",
        kind: "textarea",
        label: "問6. 復旧期の理想マップと実際マップを比較し、最も重要と思う構造的差異を1つ挙げて説明せよ。（100字以内）",
        placeholder: "例）理想マップでは…が存在するが、実際マップでは…",
        maxLength: 100,
      },
      {
        id: "q7",
        kind: "textarea",
        label: "問7. 問6で指摘した差異が生じた理由を、ICS原則または急性期との継続性の観点から説明せよ。（200字以内）",
        placeholder: "例）急性期では…であったが、復旧期には…",
        maxLength: 200,
      }
    ]
  };

  // ================================================================
  // RECOVERY_RECORD_CONTENT — 対応検証記録フェーズのデータ定義（復旧期）
  // ================================================================
  const RECOVERY_RECORD_CONTENT = {
    title: "対応検証記録（復旧期 課題抜粋）",
    excerpts: [
      {
        id: "2-1",
        text: "DHEATの撤退判断が地域の実情を踏まえず一律に行われたため、撤退後に保健所の調整機能が低下し、支援チームの活動が散漫になった。撤退時期の判断基準について、地域側との事前合意が不足していた。"
      },
      {
        id: "2-2",
        text: "復旧期においても、県庁保健医療調整本部と現地保健所との情報共有は不十分であった。仮設住宅の建設スケジュールや入居者の健康状態に関する情報が、現地から県庁に届かないケースが続いた。"
      },
      {
        id: "2-3",
        text: "在宅避難者への訪問支援の窓口が、地域包括支援センター・市町村保健センター・DWAT等で重複していた。それぞれが独自に動いており、支援の優先順位や役割分担について統一した調整の場が設けられていなかった。"
      },
      {
        id: "2-4",
        text: "福祉避難所への保健医療支援については、一般避難所と異なる対応が必要にもかかわらず、支援チームの配置基準や巡回スケジュールが一般避難所と同一であった。担当チームが福祉避難所の特性を十分に把握していなかった。"
      },
      {
        id: "2-5",
        text: "急性期から復旧期への移行期に支援チームの引き継ぎが不十分で、急性期に構築された連携関係が復旧期に継続されなかった。特にDPATとDWATの活動情報が後続の支援チームに共有されていなかった。"
      }
    ],
    questions: [
      {
        id: "q8",
        kind: "singleChoice",
        label: "問8．問7で指摘したICS原則違反は，復旧期対応検証記録の何番と対応するか1つ示せ。",
        optionsSource: "excerpts",
        required: true
      },
      {
        id: "q9",
        kind: "textarea",
        label: "問9．問6・問7での指摘を踏まえつつ、復旧期の組織構造上の問題が対応失敗をどのように引き起こしたか説明せよ。（200字以内）",
        placeholder: "例）復旧期には…という構造的問題があったため、…という失敗が生じた。",
        maxLength: 200,
        required: true
      }
    ]
  };

  // === レイヤー定義 ===
  // layer は学習者に見せる ICS 指揮階層（UI の中心概念）
  const LAYER_NAMES = ["", "指揮（Command）", "調整・統制（Section）", "実働（Branch/Group）", "支援対象"];

  // group は学習者に見せる ICS 階層ではなく内部メタデータ（将来の採点・バリデーション用）
  // UI / 描画クラスには一切使用しない
  const GROUP_EXPECTED_LAYERS = {
    "g-command": [1, 2], // 県庁・保健所・コーディネーター
    "g-section": [2],    // DHEAT のみ（調整・統制機能）
    "g-unit":    [2, 3], // 実働チーム・専門職団体
    "g-team":    [4],    // 支援対象（beneficiary）
  };

  // ================================================================
  // STATE
  // ================================================================
  const state = {
    phase: 0,
    nodes: [],
    edges: [],
    selectedNodeId: null,
    selectedEdgeId: null,
    highlightNodeId: null,
    // arrow drawing
    arrowFrom: null,   // node id
    drawingArrow: false,
    previewEnd: { x: 0, y: 0 },
    // answers
    answers: { q1: "", q2: "" },
    log: [],
  };

  window.idealMapAcute     = null;
  window.actualMapAcute    = null;
  window.actualMapRecovery = null;
  const mapLoadStatus = {
    idealAcute:     "idle",
    actualAcute:    "idle",
    actualRecovery: "idle",
  };
  window.phase5Data       = { removals: [] };

  // ── Phase6 状態変数 ────────────────────────────────────────────────────
  // Phase6 の状態は以下の 2 フラグ + phaseData.p6 の 3 点で表現する。
  //
  //   phase6Initialized      : initPhase6Canvas() が完了して phaseData.p6 に
  //                            初期ノードが入っている場合 true。
  //                            false は「未初期化」または「無効化済み（stale）」。
  //
  //   phase6RemovalSignature : 直近の initPhase6Canvas() / importJSON 時点での
  //                            削除候補スナップショット（getRemovalSignature() の値）。
  //                            switchPhase(RECOVERY_MAP) → p6NeedsRebuild() が
  //                            このシグネチャと現在値を比較して再構築の要否を判定する。
  //                            phase6Initialized=false のときは意味を持たない。
  //
  //   phaseData.p6           : Phase6 キャンバスの実データ。
  //                            phase6Initialized=true のときのみ有効な内容を持つ。
  //
  // 状態遷移:
  //   初期 / resetAll / importJSON(v1)  →  未初期化（false / ""）
  //   initPhase6Canvas() 実行後         →  有効（true / 現在のシグネチャ）
  //   toggleRemovalCandidate() 後       →  無効化（false / "" / p6 cleared）
  //   importJSON(v2) 読込後             →  有効（インポートした p6 + 復元シグネチャ）
  let   phase6Initialized      = false;
  let   phase6RemovalSignature = "";

  // ================================================================
  // PHASE DATA STORE
  // ================================================================
  const phaseData = {
    acute:    { nodes:[], edges:[], answers:{q1:"",q2:"",p3q1:"",p3q2:"",p3q2sel:""}, log:[],
                selectedNodeId:null, selectedEdgeId:null },
    // 旧フォーマット (v1-v3) の importJSON 後方互換のためのみ保持。
    // 通常フローでは savePhaseData の対象外で、常に空のまま。
    recovery: { nodes:[], edges:[], answers:{q1:"",q2:""}, log:[],
                selectedNodeId:null, selectedEdgeId:null },
    // answers.q1, q2 は現行フローでは未使用だが、importJSON v3 互換で復元先として保持。
    p6:       { nodes:[], edges:[], answers:{q1:"",q2:""}, log:[],
                selectedNodeId:null, selectedEdgeId:null },
    acuteRecord:     { answers: { q4: "", q5: "" } },
    recoveryCompare: { answers: { q6: "", q7: "", q7sel: "" } },
    recoveryRecord:  { answers: { q8: "", q9: "" } },
  };

  function savePhaseData(key) {
    phaseData[key] = {
      nodes: state.nodes, edges: state.edges,
      answers: { ...state.answers }, log: state.log,
      selectedNodeId: state.selectedNodeId, selectedEdgeId: state.selectedEdgeId,
    };
    saveToLocalStorage();
  }

  function loadPhaseData(key) {
    const d = phaseData[key];
    state.nodes = d.nodes; state.edges = d.edges;
    state.answers = { ...d.answers }; state.log = d.log;
    state.selectedNodeId = d.selectedNodeId; state.selectedEdgeId = d.selectedEdgeId;
  }

  // ================================================================
  // DOM REFS
  // ================================================================
  const $ = id => document.getElementById(id);
  const phaseSteps = document.querySelectorAll(".phase-step");

  // PHASE_VIEW_ID マップを使い、DOM 順に依存せずビューを切り替える
  function activatePhaseView(p) {
    document.querySelectorAll(".phase-view").forEach(el => el.classList.remove("active"));
    $(PHASE_VIEW_ID[p])?.classList.add("active");
  }

  // ヘッダーステップの active / done クラスを data-phase 属性で更新する
  // DOM 順の index ではなく各ステップが持つ data-phase 値で比較するため、
  // ステップの並び順変更に強い。
  function updatePhaseSteps(p) {
    phaseSteps.forEach(s => {
      const sp = parseInt(s.dataset.phase, 10);
      s.classList.toggle("active", sp === p);
      s.classList.toggle("done",   sp < p);
    });
  }

  let canvasEl           = null;
  let svgEl              = null;
  let paletteEl          = null;
  let canvasWrap         = null;
  let activeCanvasStatEl = null;
  let activeArrowHintEl  = null;
  let activeMarkerSuffix = "-acute";
  let BENEFICIARY_LABELS = new Set(["避難所", "医療機関"]);
  let activePaletteNodes = PALETTE_NODES;
  let activePhaseKey     = null;
  let _clickTimer        = null;  // シングル/ダブルクリック判定用タイマー

  // ================================================================
  // UTILS
  // ================================================================
  const uid   = () => Math.random().toString(36).slice(2, 9);
  const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
  const esc   = s => String(s)
    .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;").replace(/'/g,"&#039;");

  function logOp(type, detail) {
    state.log.push({ ts: new Date().toISOString(), type, detail });
  }

  function hasUnsavedWork() {
    const hasString = obj => obj && Object.values(obj).some(v => typeof v === "string" && v.length > 0);
    if (phaseData.acute.nodes.length > 0)    return true;
    if (phaseData.acute.edges.length > 0)    return true;
    if (phaseData.p6.nodes.length > 0)       return true;
    if (phaseData.p6.edges.length > 0)       return true;
    if (hasString(phaseData.acute.answers))           return true;
    if (hasString(phaseData.recoveryCompare?.answers)) return true;
    if (hasString(phaseData.acuteRecord?.answers))     return true;
    if (hasString(phaseData.recoveryRecord?.answers))  return true;
    if ((window.phase5Data?.removals?.length ?? 0) > 0) return true;
    return false;
  }

  // ================================================================
  // LOCAL STORAGE — 自動保存・復元
  // ================================================================
  function saveToLocalStorage() {
    try {
      const payload = {
        version:      STORAGE_SCHEMA_VERSION,
        savedAt:      new Date().toISOString(),
        currentPhase: state.phase,
        phaseData:    phaseData,
        phase5Data:   window.phase5Data,
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch (e) {
      console.warn("[ICS] localStorage への保存に失敗しました:", e);
    }
  }

  function loadFromLocalStorage() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      const obj = JSON.parse(raw);
      if (obj.version !== STORAGE_SCHEMA_VERSION) {
        // バージョン不一致：将来的にマイグレーション処理を追加する場所
        console.warn("[ICS] localStorage のスキーマバージョンが異なります。復元をスキップします。");
        return null;
      }
      return obj;
    } catch (e) {
      console.warn("[ICS] localStorage の読み込みに失敗しました:", e);
      return null;
    }
  }

  function clearLocalStorage() {
    localStorage.removeItem(STORAGE_KEY);
  }

  let _autoSaveTimer = null;
  function debouncedSave() {
    clearTimeout(_autoSaveTimer);
    _autoSaveTimer = setTimeout(saveToLocalStorage, 700);
  }

  function getNodeEl(id) {
    return canvasEl ? canvasEl.querySelector(`.node[data-id="${id}"]`) : null;
  }

  // ================================================================
  // PHASE SWITCHING
  // ================================================================
  window.switchPhase = function(p) {
    const prevPhase = state.phase;

    // drawingArrow が残留していたら必ずキャンセル
    if (state.drawingArrow) cancelArrowDraw();

    // 比較・分析画面を離れる際にフィット用 transform をリセット
    if (state.phase === PHASE.ACUTE_COMPARE) {
      for (const id of ["canvas-ideal", "svgLayer-ideal", "canvas-actual", "svgLayer-actual"]) {
        const el = $(id);
        if (el) { el.style.transform = ""; el.style.transformOrigin = ""; }
      }
    }
    if (state.phase === PHASE.RECOVERY_COMPARE) {
      for (const id of ["canvas-rcIdeal", "svgLayer-rcIdeal", "canvas-rcActual", "svgLayer-rcActual"]) {
        const el = $(id);
        if (el) { el.style.transform = ""; el.style.transformOrigin = ""; }
      }
    }

    // 現フェーズがマップ画面なら矢印キャンセル＋保存（読み取り専用フェーズは保存不要）
    if (MAP_PHASE_CONFIG[state.phase] && !MAP_PHASE_CONFIG[state.phase].isReadOnly) {
      if (state.drawingArrow) cancelArrowDraw();
      savePhaseData(activePhaseKey);
    }

    state.phase = p;
    // DOM 順ではなく PHASE_VIEW_ID マップでビューを切り替える
    activatePhaseView(p);
    // ヘッダーステップは data-phase 属性で比較する（DOM 順非依存）
    updatePhaseSteps(p);

    // ── 比較・分析（急性期） ─────────────────────────────────────────
    if (p === PHASE.ACUTE_COMPARE) {
      // 未作成チェック
      if (phaseData.acute.nodes.length === 0) {
        showToast("先に急性期の理想マップを作成してください", 3000);
        state.phase = prevPhase;
        activatePhaseView(prevPhase);
        updatePhaseSteps(prevPhase);
        return;
      }

      // 描画はグリッドレイアウト確定後に実行
      requestAnimationFrame(() => {
        // 左カラム：学習者の理想マップを描画
        renderReadOnlyMap(
          phaseData.acute.nodes,
          phaseData.acute.edges,
          $("canvas-ideal"),
          $("svgLayer-ideal"),
          $("canvasWrap-ideal"),
          $("canvasStat-ideal"),
          "-ideal",
          null, true,
          () => clearHighlightRO($("canvas-actual"), $("svgLayer-actual"))
        );

        // 右カラム：急性期実際マップを描画
        const actualCanvas = $("canvas-actual");
        const actualSvg    = $("svgLayer-actual");
        const actualWrap   = $("canvasWrap-actual");
        const actualStat   = $("canvasStat-actual");

        if (mapLoadStatus.actualAcute === "ready") {
          renderReadOnlyMap(
            window.actualMapAcute.nodes,
            window.actualMapAcute.edges,
            actualCanvas, actualSvg, actualWrap, actualStat, "-actual",
            null, true,
            () => clearHighlightRO($("canvas-ideal"), $("svgLayer-ideal"))
          );
        } else if (mapLoadStatus.actualAcute === "error") {
          actualCanvas.innerHTML =
            '<div style="color:var(--red);padding:20px;font-size:14px;">⚠ 実際マップの読み込みに失敗しました</div>';
        } else {
          actualCanvas.innerHTML =
            '<div style="color:var(--text-dim);padding:20px;font-size:14px;">読み込み中…</div>';
        }
      });

      // 回答の復元
      const ans  = phaseData.acute.answers;
      const p3q1 = $("p3q1Answer");
      const p3q2 = $("p3q2Answer");
      if (p3q1) {
        p3q1.value = ans.p3q1 || "";
        $("p3q1CharCount").textContent = (ans.p3q1 || "").length;
      }
      if (p3q2) {
        p3q2.value = ans.p3q2 || "";
        $("p3q2CharCount").textContent = (ans.p3q2 || "").length;
      }
      if (ans.p3q2sel) {
        const radio = document.querySelector(
          `input[name="p3q2principle"][value="${ans.p3q2sel}"]`
        );
        if (radio) radio.checked = true;
      }
      showToast("ノードをダブルクリックすると接続関係をハイライトできます", 3500);
      return;
    }

    // ── 対応検証記録（急性期） ─────────────────────────────────────────── [ADDED]
    if (p === PHASE.ACUTE_RECORD) {
      // 前提チェック：ACUTE_COMPARE が未完了なら戻す
      if (phaseData.acute.nodes.length === 0) {
        showToast("先に急性期の比較・分析を完了してください", 3000);
        state.phase = prevPhase;
        activatePhaseView(prevPhase);
        updatePhaseSteps(prevPhase);
        return;
      }
      renderAcuteRecordView();
      restoreAcuteRecordAnswers();
      renderSelectedPrinciple("arSelectedPrinciple", phaseData.acute.answers.p3q2sel, "問3で原則を選択すると表示されます");
      return;
    }

    // ── 復旧期比較・分析 ─────────────────────────────────────────────────── [ADDED]
    if (p === PHASE.RECOVERY_COMPARE) {
      if (phaseData.p6.nodes.length === 0) {
        showToast("先に復旧期マップを作成してください", 3000);
        state.phase = prevPhase;
        activatePhaseView(prevPhase);
        updatePhaseSteps(prevPhase);
        return;
      }
      BENEFICIARY_LABELS = PHASE6_BENEFICIARY_LABELS;

      requestAnimationFrame(() => {
        // 左カラム：学習者の復旧期理想マップ
        renderReadOnlyMap(
          phaseData.p6.nodes,
          phaseData.p6.edges,
          $("canvas-rcIdeal"),
          $("svgLayer-rcIdeal"),
          $("canvasWrap-rcIdeal"),
          $("canvasStat-rcIdeal"),
          "-rcIdeal",
          null, true,
          () => clearHighlightRO($("canvas-rcActual"), $("svgLayer-rcActual"))
        );

        // 右カラム：復旧期実際マップ
        const rcActualCanvas = $("canvas-rcActual");
        if (mapLoadStatus.actualRecovery === "ready") {
          const rcNodes = window.actualMapRecovery.recovery?.nodes || [];
          const rcEdges = window.actualMapRecovery.recovery?.edges || [];
          renderReadOnlyMap(
            rcNodes, rcEdges,
            rcActualCanvas,
            $("svgLayer-rcActual"),
            $("canvasWrap-rcActual"),
            $("canvasStat-rcActual"),
            "-rcActual",
            null, true,
            () => clearHighlightRO($("canvas-rcIdeal"), $("svgLayer-rcIdeal"))
          );
        } else if (mapLoadStatus.actualRecovery === "error") {
          if (rcActualCanvas) rcActualCanvas.innerHTML =
            '<div style="color:var(--red);padding:20px;font-size:14px;">⚠ 実際マップの読み込みに失敗しました</div>';
        } else {
          if (rcActualCanvas) rcActualCanvas.innerHTML =
            '<div style="color:var(--text-dim);padding:20px;font-size:14px;">読み込み中…</div>';
        }
      });

      // 回答の復元
      const rcAns = phaseData.recoveryCompare.answers;
      const q6el = $("rcQ6Answer");
      const q7el = $("rcQ7Answer");
      if (q6el) {
        q6el.value = rcAns.q6 || "";
        const cc = $("rcQ6CharCount");
        if (cc) cc.textContent = (rcAns.q6 || "").length;
      }
      if (q7el) {
        q7el.value = rcAns.q7 || "";
        const cc = $("rcQ7CharCount");
        if (cc) cc.textContent = (rcAns.q7 || "").length;
      }
      if (rcAns.q7sel) {
        const radio = document.querySelector(`input[name="rcQ7principle"][value="${rcAns.q7sel}"]`);
        if (radio) radio.checked = true;
      }
      showToast("ノードをダブルクリックすると接続関係をハイライトできます", 3500);
      return;
    }

    // ── 対応検証記録（復旧期） ────────────────────────────────────────────
    if (p === PHASE.RECOVERY_RECORD) {
      if (phaseData.p6.nodes.length === 0) {
        showToast("先に復旧期マップを作成してください", 3000);
        state.phase = prevPhase;
        activatePhaseView(prevPhase);
        updatePhaseSteps(prevPhase);
        return;
      }
      renderRecoveryRecordView();
      restoreRecoveryRecordAnswers();
      renderSelectedPrinciple("rrSelectedPrinciple", phaseData.recoveryCompare.answers.q7sel, "問7で原則を選択すると表示されます");
      return;
    }

    // ── マップ系フェーズ共通（急性期マップ・復旧期準備・復旧期マップ） ───
    const cfg = MAP_PHASE_CONFIG[p];
    if (cfg) {
      if (cfg.isReadOnly) {
        // 復旧期準備：読み取り専用モード
        activePhaseKey     = cfg.key;
        BENEFICIARY_LABELS = cfg.beneficiaries;
        canvasEl           = $(cfg.domIds.canvas);
        svgEl              = $(cfg.domIds.svg);
        canvasWrap         = $(cfg.domIds.wrap);
        activeCanvasStatEl = $(cfg.domIds.stat);
        activeArrowHintEl  = null;
        activeMarkerSuffix = cfg.markerSuffix;
        activePaletteNodes = [];
        renderPhase5Map();
      } else {
        // 通常モード（急性期マップ・復旧期マップ）
        activePhaseKey     = cfg.key;
        activePaletteNodes = cfg.paletteNodes;
        BENEFICIARY_LABELS = cfg.beneficiaries;
        canvasEl           = $(cfg.domIds.canvas);
        svgEl              = $(cfg.domIds.svg);
        paletteEl          = $(cfg.domIds.palette);
        canvasWrap         = $(cfg.domIds.wrap);
        activeCanvasStatEl = $(cfg.domIds.stat);
        activeArrowHintEl  = $(cfg.domIds.hint);
        activeMarkerSuffix = cfg.markerSuffix;

        if (p === PHASE.RECOVERY_MAP) {
          // ヘッダーからの直接遷移も含め、削除候補が未選択なら復旧期準備に誘導する
          if (window.phase5Data.removals.length === 0) {
            showToast("先に復旧期準備で不要なノードを選択してください", 3000);
            state.phase = prevPhase;
            activatePhaseView(prevPhase);
            updatePhaseSteps(prevPhase);
            return;
          }
          // 未初期化、または削除候補が変わった場合に再構築する。
          // p6NeedsRebuild() は phase6Initialized と phase6RemovalSignature の両方で判定するため、
          // toggleRemovalCandidate() の直後に遷移してきた場合も確実に再構築される。
          if (p6NeedsRebuild()) {
            initPhase6Canvas();
          }
        }
        loadPhaseData(cfg.key);
        renderPalette();
        renderAll();
      }
    }
  };
  // data-phase 属性の値を使うことで、DOM順の index に依存しない
  phaseSteps.forEach(s => s.addEventListener("click", () => switchPhase(parseInt(s.dataset.phase, 10))));
  $("btnStartMap").addEventListener("click", () => switchPhase(PHASE.ACUTE_MAP));

  // ================================================================
  // PALETTE
  // ================================================================
  function renderPalette() {
    paletteEl.innerHTML = "";

    // Phase6 の場合：キャンバス配置済みラベルの集合を作成
    const placedLabels = (activePhaseKey === "p6")
      ? new Set(state.nodes.map(n => n.label))
      : null;

    for (const n of activePaletteNodes) {
      const div = document.createElement("div");
      // group は内部メタデータ。パレット item に group クラスを付与しない（layer のみで色を表現）
      div.className = "pitem";
      if (BENEFICIARY_LABELS.has(n.label)) div.classList.add("node-beneficiary");

      const isPlaced = placedLabels?.has(n.label) ?? false;
      if (isPlaced) div.classList.add("pitem-placed");

      div.innerHTML = `
        <span class="pico">${n.icon || ""}</span>
        <span class="plabel">${esc(n.label)}</span>
        <span class="ptag">${isPlaced ? "配置済み" : "＋追加"}</span>
      `;

      if (!isPlaced) {
        div.addEventListener("click", () => {
          addNode(n.label, n.group);
          logOp("ADD_NODE", { label: n.label, group: n.group });
        });
      }
      paletteEl.appendChild(div);
    }
  }

  // ================================================================
  // NODE OPERATIONS
  // ================================================================
  function addNode(label, group) {
    const rect = canvasWrap.getBoundingClientRect();
    const x = 80 + Math.floor(Math.random() * Math.max(rect.width - 280, 80));
    const y = 60 + Math.floor(Math.random() * Math.max(rect.height - 120, 60));
    const id = "n-" + uid();
    state.nodes.push({ id, label, group, x, y, layerId: BENEFICIARY_LABELS.has(label) ? 4 : null, layerReason: "" });
    renderAll();
    selectNode(id);
    saveToLocalStorage();
  }

  function selectNode(id) {
    if (state.selectedNodeId === id) {
      clearSelection();
      return;
    }
    state.selectedNodeId = id;
    state.selectedEdgeId = null;
    const n = state.nodes.find(x => x.id === id);
    if (!n) return;
    canvasEl.querySelectorAll(".node").forEach(el => el.classList.remove("selected"));
    getNodeEl(id)?.classList.add("selected");
    // 削除ボタンの表示切替
    canvasEl.querySelectorAll(".node-delete-btn").forEach(b => b.style.display = "none");
    const btn = getNodeEl(id)?.querySelector(".node-delete-btn");
    if (btn) btn.style.display = "flex";
    updateQ1Select();
  }

  function clearSelection() {
    state.selectedNodeId = null;
    clearHighlight();
    canvasEl?.querySelectorAll(".node").forEach(el => el.classList.remove("selected"));
    canvasEl?.querySelectorAll(".node-delete-btn").forEach(b => b.style.display = "none");
  }

  function applyHighlight(selectedId) {
    const connected = new Set();
    for (const e of state.edges) {
      if (e.from === selectedId) connected.add(e.to);
      if (e.to   === selectedId) connected.add(e.from);
    }
    canvasEl.querySelectorAll(".node").forEach(el => {
      el.classList.remove("node-focus", "node-active", "node-dim");
      const nid = el.dataset.id;
      if (nid === selectedId)      el.classList.add("node-focus");
      else if (connected.has(nid)) el.classList.add("node-active");
      else                         el.classList.add("node-dim");
    });
    svgEl.querySelectorAll("g[data-from]").forEach(g => {
      const isActive = g.dataset.from === selectedId || g.dataset.to === selectedId;
      g.setAttribute("opacity", isActive ? "1" : "0.1");
      if (isActive) {
        g.style.filter = "drop-shadow(0 0 4px rgba(255,255,255,0.35))";
      }
    });
  }

  function clearHighlight() {
    state.highlightNodeId = null;
    canvasEl?.querySelectorAll(".node").forEach(el => {
      el.classList.remove("node-focus", "node-active", "node-dim");
    });
    svgEl?.querySelectorAll("g[data-from]").forEach(g => {
      g.setAttribute("opacity", "1");
      g.style.filter = "";
    });
  }

  function applyHighlightRO(selectedId, nodes, edges, panelCanvas, panelSvg) {
    const connected = new Set();
    for (const e of edges) {
      if (e.from === selectedId) connected.add(e.to);
      if (e.to   === selectedId) connected.add(e.from);
    }
    panelCanvas.querySelectorAll(".node").forEach(el => {
      el.classList.remove("node-focus", "node-active", "node-dim");
      const nid = el.dataset.id;
      if      (nid === selectedId)  el.classList.add("node-focus");
      else if (connected.has(nid))  el.classList.add("node-active");
      else                          el.classList.add("node-dim");
    });
    panelSvg.querySelectorAll("g[data-from]").forEach(g => {
      const active = g.dataset.from === selectedId || g.dataset.to === selectedId;
      g.setAttribute("opacity", active ? "1" : "0.1");
      g.style.filter = active
        ? "drop-shadow(0 0 4px rgba(255,255,255,0.35))"
        : "";
    });
  }

  function clearHighlightRO(panelCanvas, panelSvg) {
    panelCanvas?.querySelectorAll(".node")
      .forEach(el => el.classList.remove("node-focus", "node-active", "node-dim"));
    panelSvg?.querySelectorAll("g[data-from]").forEach(g => {
      g.setAttribute("opacity", "1");
      g.style.filter = "";
    });
  }

  // Phase5 の削除候補リストを一意なシグネチャ文字列に変換する。
  // nodeId をソートして結合するため、同一セットなら挿入順によらず同じ値になる。
  function getRemovalSignature() {
    return window.phase5Data.removals
      .map(r => r.nodeId)
      .sort()
      .join(",");
  }

  // Phase6 に「ユーザーが加えた編集」が存在するかを判定する。
  // isInitial=false のノード（自分で追加したもの）または矢印が1本以上あれば編集済み。
  // toggleRemovalCandidate() で confirm を出すかどうかの判断に使う。
  function hasP6Edits() {
    return phaseData.p6.nodes.some(n => !n.isInitial) || phaseData.p6.edges.length > 0;
  }

  // Phase6 の再構築（initPhase6Canvas の再実行）が必要かどうかを判定する。
  // ・未初期化（phase6Initialized=false）
  // ・削除候補が変わった（現在のシグネチャ ≠ 構築時のシグネチャ）
  // のどちらかで true を返す。switchPhase(RECOVERY_MAP) はこの関数でのみ再構築の要否を判断する。
  function p6NeedsRebuild() {
    return !phase6Initialized || getRemovalSignature() !== phase6RemovalSignature;
  }

  // Phase6 を「無効化済み」状態にセットする。
  // 呼び出し後は p6NeedsRebuild()=true となり、次回 Phase6 入場時に initPhase6Canvas() が走る。
  // phaseData.p6 も即時クリアして stale データの誤利用を防ぐ。
  // この関数が Phase6 の「初期化フラグ + シグネチャ + データ」をまとめてリセットする唯一の責務を持つ。
  function invalidatePhase6() {
    phase6Initialized      = false;
    phase6RemovalSignature = "";
    phaseData.p6 = {
      nodes: [], edges: [], answers: { q1: "", q2: "" },
      log: [], selectedNodeId: null, selectedEdgeId: null,
    };
  }

  function deleteNode(id) {
    const n = state.nodes.find(x => x.id === id);
    if (n?.isInitial) {
      showToast("急性期理想マップ由来のノードは削除できません");
      return;
    }
    if (!confirm(`「${n?.label}」を削除しますか？`)) return;
    logOp("DELETE_NODE", { id, label: n?.label });
    state.edges = state.edges.filter(e => e.from !== id && e.to !== id);
    state.nodes = state.nodes.filter(n => n.id !== id);
    state.selectedNodeId = null;
    renderAll();
    saveToLocalStorage();
  }

  // ================================================================
  // NODE DOM RENDER
  // ================================================================
  function renderNodes() {
    canvasEl.innerHTML = "";
    const cfg = MAP_PHASE_CONFIG[state.phase];
    const isReadOnly = !!(cfg?.isReadOnly);

    for (const n of state.nodes) {
      const div = document.createElement("div");
      const layerClass = n.layerId ? `layer-${n.layerId}` : "layer-none";
      const benefClass = BENEFICIARY_LABELS.has(n.label) ? " node-beneficiary" : "";
      // group は内部メタデータ。class には layer と beneficiary のみ反映する
      div.className = `node ${isReadOnly ? "" : "draggable "}${layerClass}${benefClass}`;
      div.dataset.id = n.id;
      div.style.left = n.x + "px";
      div.style.top  = n.y + "px";
      if (n.id === state.selectedNodeId) div.classList.add("selected");

      div.innerHTML = `<div class="ntitle">${esc(n.label)}</div>`;

      if (!isReadOnly) {
        div.innerHTML += `<div class="node-connect-btn" data-nid="${n.id}" title="矢印を引く">→</div>`;

        // isInitial ノードには削除ボタンを付けない
        if (!n.isInitial) {
          const deleteBtn = document.createElement("div");
          deleteBtn.className = "node-delete-btn";
          deleteBtn.textContent = "×";
          deleteBtn.style.display = n.id === state.selectedNodeId ? "flex" : "none";
          deleteBtn.addEventListener("click", e => {
            e.stopPropagation();
            deleteNode(n.id);
          });
          div.appendChild(deleteBtn);
        }
      }

      // 削除候補クラス（Phase5）
      if (isReadOnly && window.phase5Data.removals.some(r => r.nodeId === n.id)) {
        div.classList.add("node-removal-candidate");
      }

      // 急性期引き継ぎノード（Phase6）
      if (n.isInitial) {
        div.classList.add("node-initial");
      }

      // Click on node div
      div.addEventListener("click", e => {
        e.stopPropagation();

        // 読み取り専用フェーズのクリックはコールバックで処理
        if (isReadOnly) {
          toggleRemovalCandidate(n.id, n.label);
          return;
        }

        // connect-btn / delete-btn のクリックは無視
        if (e.target.classList.contains("node-connect-btn")) return;
        if (e.target.classList.contains("node-delete-btn")) return;

        // 矢印描画モード中はターゲット選択として処理
        if (state.drawingArrow) {
          if (state.arrowFrom !== n.id) {
            finishArrow(n.id, e.clientX, e.clientY);
          }
          return;
        }

        // 通常クリック
        clearTimeout(_clickTimer);
        _clickTimer = setTimeout(() => { selectNode(n.id); }, 250);
      });

      // Double-click: ハイライトモード発動 / 解除
      div.addEventListener("dblclick", e => {
        e.stopPropagation();
        if (isReadOnly) return;
        if (e.target.classList.contains("node-connect-btn")) return;
        if (e.target.classList.contains("node-delete-btn")) return;
        clearTimeout(_clickTimer);
        if (state.highlightNodeId === n.id) {
          clearHighlight();
        } else {
          state.highlightNodeId = n.id;
          applyHighlight(n.id);
        }
      });

      // Arrow hover highlight while drawing
      if (!isReadOnly) {
        div.addEventListener("mouseenter", () => {
          if (state.drawingArrow && state.arrowFrom !== n.id) {
            div.classList.add("arrow-target-hover");
          }
        });
        div.addEventListener("mouseleave", () => {
          div.classList.remove("arrow-target-hover");
        });

        // Drag to move
        setupDrag(div, n);
      }

      canvasEl.appendChild(div);
    }

    // Connect buttons: stop ALL propagation so click never reaches parent node div
    if (!isReadOnly) {
      canvasEl.querySelectorAll(".node-connect-btn").forEach(btn => {
        btn.addEventListener("mousedown", e => { e.stopPropagation(); e.preventDefault(); });
        btn.addEventListener("click", e => {
          e.stopPropagation();
          e.preventDefault();
          startArrowDraw(btn.dataset.nid);
        });
      });
    }

    // Canvas background click: only deselect; never cancel arrow (Esc key does that)
    // canvasEl は pointer-events:none のため、空白クリックは canvasWrap に届く
    canvasWrap.onclick = e => {
      if (isReadOnly) return;
      if (state.drawingArrow) return;
      if (e.target !== canvasWrap) return;
      clearSelection();
      if (state.selectedEdgeId) { state.selectedEdgeId = null; renderEdges(); }
    };
  }

  // ================================================================
  // DRAG-TO-MOVE
  // ================================================================
  function setupDrag(div, n) {
    div.addEventListener("pointerdown", e => {
      if (state.drawingArrow) return;
      if (e.target.classList.contains("node-connect-btn")) return;
      if (e.target.classList.contains("node-delete-btn")) return;
      e.preventDefault();
      div.setPointerCapture(e.pointerId);
      // ドラッグ中は削除ボタンを非表示（誤タップ防止）
      const delBtn = div.querySelector(".node-delete-btn");
      if (delBtn) delBtn.style.display = "none";
      const sx = e.clientX, sy = e.clientY, bx = n.x, by = n.y;
      let moved = false;

      const onMove = ev => {
        moved = true;
        const rect = canvasWrap.getBoundingClientRect();
        n.x = clamp(bx + ev.clientX - sx, 0, rect.width - 200);
        n.y = clamp(by + ev.clientY - sy, 0, rect.height - 80);
        div.style.left = n.x + "px";
        div.style.top  = n.y + "px";
        highlightLayer(getLayerIdFromY(n.y));
        renderEdges();
        updateCanvasStat();
      };
      const onUp = () => {
        div.removeEventListener("pointermove", onMove);
        div.removeEventListener("pointerup", onUp);
        div.releasePointerCapture(e.pointerId);
        clearLayerHighlight();
        if (moved) {
          if (!BENEFICIARY_LABELS.has(n.label)) {
            n.layerId = getLayerIdFromY(n.y);
            // 層クラスをノード要素に即時反映
            div.classList.remove("layer-none", "layer-1", "layer-2", "layer-3", "layer-4");
            div.classList.add(`layer-${n.layerId}`);
          }
          logOp("MOVE_NODE", { id: n.id, label: n.label, layerId: n.layerId });
          saveToLocalStorage();
        }
      };
      div.addEventListener("pointermove", onMove);
      div.addEventListener("pointerup", onUp);
    });
  }

  // ================================================================
  // LAYER HELPERS
  // ================================================================
  function getLayerIdFromY(y) {
    const rect = canvasWrap.getBoundingClientRect();
    const pct = y / rect.height;
    if (pct < 0.25) return 1;
    if (pct < 0.50) return 2;
    if (pct < 0.75) return 3;
    return 4;
  }

  function highlightLayer(layerId) {
    canvasWrap.querySelectorAll(".layer-band").forEach(b => {
      b.classList.toggle("active", +b.dataset.layer === layerId);
    });
  }

  function clearLayerHighlight() {
    canvasWrap.querySelectorAll(".layer-band").forEach(b => b.classList.remove("active"));
  }

  // ================================================================
  // ARROW DRAWING — PowerPoint style
  // ================================================================
  function startArrowDraw(fromId) {
    clearHighlight();
    const fromNode = state.nodes.find(n => n.id === fromId);
    if (BENEFICIARY_LABELS.has(fromNode?.label)) return; // 被支援者ノードからは矢印不可
    state.arrowFrom = fromId;
    state.drawingArrow = true;
    canvasWrap.classList.add("drawing-arrow");

    // highlight source node
    canvasEl.querySelectorAll(".node").forEach(el => el.classList.remove("arrow-source"));
    getNodeEl(fromId)?.classList.add("arrow-source");

    // show hint
    if (activeArrowHintEl) activeArrowHintEl.style.display = "flex";

    // mousemove on canvas for rubber-band preview
    canvasWrap.addEventListener("mousemove", onArrowMouseMove);
    // escape key cancel
    document.addEventListener("keydown", onArrowKeyDown);
  }

  function onArrowMouseMove(e) {
    if (!state.drawingArrow) return;
    const rect = canvasWrap.getBoundingClientRect();
    state.previewEnd = { x: e.clientX - rect.left, y: e.clientY - rect.top };
    renderArrowPreview();
  }

  function onArrowKeyDown(e) {
    if (e.key === "Escape") cancelArrowDraw();
  }

  function finishArrow(toId, clientX, clientY) {
    if (!state.drawingArrow || !state.arrowFrom) return;
    if (state.arrowFrom === toId) { cancelArrowDraw(); return; }

    const fromId = state.arrowFrom;
    cancelArrowDraw();
    // 支援対象（BENEFICIARY_LABELS に含まれるノード）への矢印は「支援」ラベルを自動付与
    const toNode = state.nodes.find(n => n.id === toId);
    if (BENEFICIARY_LABELS.has(toNode?.label)) {
      addEdgeWithLabel(fromId, toId, "支援");
      return;
    }
    showEdgeLabelPopup(fromId, toId, clientX, clientY);
  }

  function cancelArrowDraw() {
    state.drawingArrow = false;
    state.arrowFrom = null;
    canvasWrap.classList.remove("drawing-arrow");
    canvasWrap.removeEventListener("mousemove", onArrowMouseMove);
    document.removeEventListener("keydown", onArrowKeyDown);
    canvasEl?.querySelectorAll(".node").forEach(el => {
      el.classList.remove("arrow-source", "arrow-target-hover");
    });
    clearArrowPreview();
    if (activeArrowHintEl) activeArrowHintEl.style.display = "none";
  }

  function renderArrowPreview() {
    const previewId = "arrowPreview" + activeMarkerSuffix;
    const prevMarkerId = "prev-arrow" + activeMarkerSuffix;
    let previewSvg = $(previewId);
    if (!previewSvg) {
      previewSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      previewSvg.id = previewId;
      previewSvg.setAttribute("style", "position:absolute;inset:0;width:100%;height:100%;pointer-events:none;z-index:4;overflow:visible;");
      previewSvg.innerHTML = `<defs>
        <marker id="${prevMarkerId}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse">
          <path d="M 0 0 L 10 5 L 0 10 z" fill="#ffd166" opacity="0.8"/>
        </marker></defs>`;
      canvasWrap.appendChild(previewSvg);
    }

    // clear old lines
    Array.from(previewSvg.querySelectorAll("line")).forEach(l => l.remove());

    const fromNode = state.nodes.find(n => n.id === state.arrowFrom);
    if (!fromNode) return;

    const a = nodeCenter(fromNode);
    const b = state.previewEnd;

    const dx = b.x - a.x, dy = b.y - a.y;
    const dist = Math.sqrt(dx*dx + dy*dy) || 1;
    const ex = b.x - (dx/dist)*14;
    const ey = b.y - (dy/dist)*14;

    const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
    line.setAttribute("x1", a.x); line.setAttribute("y1", a.y);
    line.setAttribute("x2", ex);  line.setAttribute("y2", ey);
    line.setAttribute("stroke", "#ffd166");
    line.setAttribute("stroke-width", "2.5");
    line.setAttribute("stroke-dasharray", "8 4");
    line.setAttribute("marker-end", `url(#prev-arrow${activeMarkerSuffix})`);
    line.setAttribute("opacity", "0.75");
    previewSvg.appendChild(line);
  }

  function clearArrowPreview() {
    const p = $("arrowPreview" + activeMarkerSuffix);
    if (p) p.querySelectorAll("line").forEach(l => l.remove());
  }

  // ================================================================
  // EDGE LABEL POPUP
  // ================================================================
  // ================================================================
  // EDGE LABEL POPUP  ― 3種類 ＋ 方向選択
  // ================================================================
  function showEdgeLabelPopup(fromId, toId, clientX, clientY) {
    document.querySelectorAll(".edge-label-popup").forEach(p => p.remove());

    const popup = document.createElement("div");
    popup.className = "edge-label-popup";
    popup.style.left = Math.min(clientX, window.innerWidth - 310) + "px";
    popup.style.top  = Math.min(clientY, window.innerHeight - 220) + "px";

    const fromNode = state.nodes.find(n => n.id === fromId);
    const toNode   = state.nodes.find(n => n.id === toId);
    const fLbl = fromNode?.label || "A";
    const tLbl = toNode?.label   || "B";

    // 方向はドラッグで確定済み。ラベル種類のみ選択させる
    const rowsHtml = EDGE_TYPES.map(t => {
      const dot = `<span class="ebdot" style="background:${t.stroke}"></span>`;
      return `
        <button class="edge-dir-btn edge-create-type-btn" data-label="${esc(t.label)}"
          style="width:100%;border-color:${t.stroke}">
          <span class="dir-txt" style="color:${t.stroke}">${dot} ${esc(t.label)}</span>
          <span class="dir-sub">${esc(t.desc)}</span>
        </button>`;
    }).join("");

    popup.innerHTML = `
      <div class="popup-title">矢印の種類を選択</div>
      <div class="popup-nodes">
        <span class="popup-node-chip from">${esc(fLbl)}</span>
        <span class="popup-node-arr">→</span>
        <span class="popup-node-chip to">${esc(tLbl)}</span>
      </div>
      <div class="popup-types" style="gap:6px">${rowsHtml}</div>
      <button class="popup-cancel">キャンセル</button>
    `;

    document.body.appendChild(popup);

    popup.querySelectorAll(".edge-create-type-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        addEdgeWithLabel(fromId, toId, btn.dataset.label);
        popup.remove();
      });
    });
    popup.querySelector(".popup-cancel").addEventListener("click", () => popup.remove());

    setTimeout(() => {
      document.addEventListener("click", function closePopup(e) {
        if (!popup.contains(e.target)) { popup.remove(); document.removeEventListener("click", closePopup); }
      });
    }, 100);
  }

  function addEdgeWithLabel(fromId, toId, label) {
    const check = canAddEdge(fromId, toId, label, state.edges);
    if (!check.allowed) { showToast(check.reason); return; }
    const id    = "e-" + uid();
    const type  = EDGE_MAP[label] || EDGE_TYPES[0];
    const bidir = !!type.bidirectional;
    state.edges.push({ id, from: fromId, to: toId, label, bidirectional: bidir });
    const fromN = state.nodes.find(n => n.id === fromId);
    const toN   = state.nodes.find(n => n.id === toId);
    logOp("ADD_EDGE", { from: fromN?.label, to: toN?.label, label, bidirectional: bidir });
    renderAll();
    saveToLocalStorage();
  }

  // ================================================================
  // EDGE RENDER (SVG)
  // ================================================================
  function nodeCenter(n) {
    const el = getNodeEl(n.id);
    if (!el) return { x: n.x + 75, y: n.y + 28 };
    const r  = el.getBoundingClientRect();
    const cr = canvasWrap.getBoundingClientRect();
    return { x: r.left - cr.left + r.width/2, y: r.top - cr.top + r.height/2 };
  }

  function renderEdges() {
    const defs = svgEl.querySelector("defs");
    svgEl.innerHTML = "";
    if (defs) svgEl.appendChild(defs);
    const isP5 = !!(MAP_PHASE_CONFIG[state.phase]?.isReadOnly);

    const CURVE_OFFSET  = 50;
    const SAME_DIR_STEP = 22; // 同方向グループ内の間隔（px）

    for (const e of state.edges) {
      const from = state.nodes.find(n => n.id === e.from);
      const to   = state.nodes.find(n => n.id === e.to);
      if (!from || !to) continue;

      const a = nodeCenter(from);
      const b = nodeCenter(to);
      const typeInfo = EDGE_MAP[e.label] || EDGE_TYPES[0];
      const col = typeInfo.stroke;

      const dx = b.x - a.x, dy = b.y - a.y;
      const dist = Math.sqrt(dx*dx + dy*dy) || 1;
      const shorten = 26;

      // 逆向きペアが存在するか確認（自身を除く）
      const hasReverse = state.edges.some(r =>
        r.id !== e.id && r.from === e.to && r.to === e.from
      );

      // 同方向グループ（同一 from/to）を ID でソートして安定インデックスを取得
      const sameDirEdges = state.edges
        .filter(x => x.from === e.from && x.to === e.to)
        .sort((p, q) => p.id < q.id ? -1 : 1);
      const sameDirIndex  = sameDirEdges.findIndex(x => x.id === e.id);
      const sameDirCount  = sameDirEdges.length;
      const sameDirOffset = (sameDirIndex - (sameDirCount - 1) / 2) * SAME_DIR_STEP;

      // 正規方向（ID小→大）で法線を統一（hasReverse の有無に関わらず常に算出）
      const sign = e.from < e.to ? 1 : -1;
      const [canonFrom, canonTo] = e.from < e.to ? [from, to] : [to, from];
      const ca = nodeCenter(canonFrom), cb = nodeCenter(canonTo);
      const cdx = cb.x - ca.x, cdy = cb.y - ca.y;
      const cdist = Math.sqrt(cdx*cdx + cdy*cdy) || 1;
      const cnx = -cdy / cdist, cny = cdx / cdist;

      // 開始・終了点（ノード端から shorten px 引く）
      const sx = a.x + (dx/dist)*shorten;
      const sy = a.y + (dy/dist)*shorten;
      const ex = b.x - (dx/dist)*shorten;
      const ey = b.y - (dy/dist)*shorten;

      let pathD, lx, ly;

      if (e.bidirectional) {
        // 双方向（連携協力）: 両端に矢印を持つ直線（変更なし）
        const bsx = a.x + (dx / dist) * shorten;
        const bsy = a.y + (dy / dist) * shorten;
        pathD = `M ${bsx} ${bsy} L ${ex} ${ey}`;
        lx = (a.x + b.x) / 2;
        ly = (a.y + b.y) / 2;
      } else {
        // 指示命令・情報伝達・支援: 複合オフセットによる二次ベジェ
        // totalOffset=0 のとき制御点が中点 → 直線と等価
        const reverseOffset = hasReverse ? sign * CURVE_OFFSET : 0;
        const totalOffset   = reverseOffset + sameDirOffset;
        const cpx = (sx + ex) / 2 + totalOffset * cnx;
        const cpy = (sy + ey) / 2 + totalOffset * cny;
        pathD = `M ${sx} ${sy} Q ${cpx} ${cpy} ${ex} ${ey}`;
        lx = (sx + 2 * cpx + ex) / 4;
        ly = (sy + 2 * cpy + ey) / 4;
      }

      const isSelected = e.id === state.selectedEdgeId;
      const markerKey = col === "#ff6b6b" ? "red" : col === "#4d8fff" ? "blue" : col === "#c084fc" ? "purple" : "teal";
      const markerEnd   = `url(#arrow-${markerKey}${activeMarkerSuffix})`;
      const markerStart = e.bidirectional ? `url(#arrow-${markerKey}${activeMarkerSuffix})` : "none";

      // <g> ラッパー（ハイライト用 data 属性付き）
      const edgeGroup = document.createElementNS("http://www.w3.org/2000/svg", "g");
      edgeGroup.dataset.from   = e.from;
      edgeGroup.dataset.to     = e.to;
      edgeGroup.dataset.edgeId = e.id;
      svgEl.appendChild(edgeGroup);

      const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
      path.setAttribute("d", pathD);
      path.setAttribute("stroke", isSelected ? "#fbbf24" : col);
      path.setAttribute("stroke-width", isSelected ? "3.5" : "2.2");
      path.setAttribute("fill", "none");
      path.setAttribute("marker-end", markerEnd);
      if (e.bidirectional) path.setAttribute("marker-start", markerStart);
      path.setAttribute("opacity", isSelected ? "1" : isP5 ? "0.4" : "0.88");
      edgeGroup.appendChild(path);

      // label pill
      const labelW = (e.label?.length || 0) * 8 + 12;
      const bg = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      bg.setAttribute("x", lx - labelW/2); bg.setAttribute("y", ly - 11);
      bg.setAttribute("width", labelW); bg.setAttribute("height", 16);
      bg.setAttribute("rx", 5); bg.setAttribute("fill", "#0d1422"); bg.setAttribute("opacity", "0.9");
      edgeGroup.appendChild(bg);

      const text = document.createElementNS("http://www.w3.org/2000/svg", "text");
      text.setAttribute("x", lx); text.setAttribute("y", ly - 2);
      text.setAttribute("class", "edge-label-text");
      text.setAttribute("fill", isSelected ? "#fbbf24" : col);
      text.textContent = e.label || "";
      edgeGroup.appendChild(text);

      // 選択中エッジに × 削除ボタンを中点に表示
      if (isSelected) {
        const delCircle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
        delCircle.setAttribute("cx", lx);
        delCircle.setAttribute("cy", ly + 14);
        delCircle.setAttribute("r", "10");
        delCircle.setAttribute("fill", "#ef4444");
        delCircle.setAttribute("cursor", "pointer");
        delCircle.setAttribute("pointer-events", "all");
        delCircle.addEventListener("click", ev => { ev.stopPropagation(); deleteEdge(e.id); });
        edgeGroup.appendChild(delCircle);

        const delText = document.createElementNS("http://www.w3.org/2000/svg", "text");
        delText.setAttribute("x", lx);
        delText.setAttribute("y", ly + 14);
        delText.setAttribute("text-anchor", "middle");
        delText.setAttribute("dominant-baseline", "central");
        delText.setAttribute("fill", "#fff");
        delText.setAttribute("font-size", "13");
        delText.setAttribute("font-weight", "900");
        delText.setAttribute("pointer-events", "none");
        delText.textContent = "×";
        edgeGroup.appendChild(delText);
      }

      // クリック用透明ヒットエリア（選択済み or 読み取り専用フェーズでは追加しない）
      if (!isSelected && !isP5) {
        const hitArea = document.createElementNS("http://www.w3.org/2000/svg", "path");
        hitArea.setAttribute("d", pathD);
        hitArea.setAttribute("stroke", "transparent");
        hitArea.setAttribute("stroke-width", "14");
        hitArea.setAttribute("fill", "none");
        hitArea.setAttribute("cursor", "pointer");
        hitArea.setAttribute("pointer-events", "all");
        hitArea.dataset.edgeId = e.id;
        hitArea.addEventListener("click", ev => {
          ev.stopPropagation();
          selectEdge(e.id);
        });
        edgeGroup.appendChild(hitArea);
      }
    }
  }

  // ================================================================
  // SELECTORS & Q1
  // ================================================================
  function updateSelectors() {
    updateQ1Select();
  }

  function updateQ1Select() {
    const sel = $("q1Answer");
    if (!sel) return;
    const prev = state.answers.q1;
    sel.innerHTML = `<option value="">（ノードを選択）</option>` +
      state.nodes.map(n => `<option value="${n.id}" ${n.id===prev?"selected":""}>${esc(n.label)}</option>`).join("");
  }

  // ================================================================
  // EDGE SELECTION / DELETION
  // ================================================================
  function selectEdge(id) {
    state.selectedEdgeId = id;
    state.selectedNodeId = null;
    document.querySelectorAll(".node").forEach(el => el.classList.remove("selected"));
    document.querySelectorAll(".node-delete-btn").forEach(b => b.style.display = "none");
    renderEdges();
  }

  function deleteEdge(id) {
    const e = state.edges.find(x => x.id === id);
    if (!e) return;
    const fromN = state.nodes.find(n => n.id === e.from);
    const toN   = state.nodes.find(n => n.id === e.to);
    if (!confirm(`「${fromN?.label} → ${toN?.label}」の矢印を削除しますか？`)) return;
    logOp("DELETE_EDGE", { from: fromN?.label, to: toN?.label, label: e.label });
    state.edges = state.edges.filter(x => x.id !== id);
    state.selectedEdgeId = null;
    renderAll();
    saveToLocalStorage();
  }

  // ================================================================
  // AUTO LAYOUT
  // ================================================================
  // ================================================================
  // CANVAS STAT
  // ================================================================
  function updateCanvasStat() {
    if (activeCanvasStatEl) activeCanvasStatEl.textContent = `ノード: ${state.nodes.length} ／ 矢印: ${state.edges.length}`;
  }

  // ================================================================
  // JSON / EXPORT
  // ================================================================
  function buildExportObject() {
    // 現フェーズの状態を一時保存（読み取り専用フェーズは保存不要）
    if (activePhaseKey && !MAP_PHASE_CONFIG[state.phase]?.isReadOnly)
      savePhaseData(activePhaseKey);
    return {
      version: 5, // [CHANGED] 4 → 5
      exportedAt: new Date().toISOString(),
      scenarioId: SCENARIO.id,
      acute: {
        nodes: phaseData.acute.nodes,
        edges: phaseData.acute.edges,
        answers: phaseData.acute.answers,
        operationLog: phaseData.acute.log,
      },
      recovery: {
        nodes: phaseData.p6.nodes,
        edges: phaseData.p6.edges,
        answers: phaseData.p6.answers,
        operationLog: phaseData.p6.log,
      },
      phase5Data: window.phase5Data,
      acuteRecord: {
        answers: phaseData.acuteRecord.answers,
      },
      recoveryCompare: {
        answers: phaseData.recoveryCompare.answers,
      },
      recoveryRecord: {
        answers: phaseData.recoveryRecord.answers,
      },
    };
  }

  function exportJSON() {
    const blob = new Blob([JSON.stringify(buildExportObject(), null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `ics_log_${SCENARIO.id}_${new Date().toISOString().slice(0,10)}.json`;
    a.click();
    URL.revokeObjectURL(a.href);
    logOp("EXPORT", {});
  }

  function validateEdgeConflicts(edges) {
    const _cmdInfoSet = new Set(["指示命令", "情報伝達"]);
    return edges.some(e1 =>
      e1.label === "連携協力" &&
      edges.some(e2 =>
        e2.id !== e1.id && _cmdInfoSet.has(e2.label) &&
        ((e2.from === e1.from && e2.to === e1.to) ||
         (e2.from === e1.to   && e2.to === e1.from))
      )
    );
  }

  function importJSON() {
    const inp = document.createElement("input");
    inp.type = "file"; inp.accept = ".json";
    inp.onchange = () => {
      const file = inp.files?.[0];
      if (!file) return;
      const fr = new FileReader();
      fr.onload = () => {
        try {
          const obj = JSON.parse(fr.result);
          let hasConflict = false;

          if ((obj.version === 5 || obj.version === 4 || obj.version === 3) && obj.acute && obj.recovery) { // v5/v4/v3
            // v3/v4: 全フェーズを復元
            const loadPhase = (src) => ({
              nodes: (src.nodes || []).map(n => ({ layerId: null, layerReason: "", isInitial: false, ...n })),
              edges: src.edges || [],
              answers: { q1: "", q2: "", p3q1: "", p3q2: "", p3q2sel: "", ...(src.answers || {}) },
              log: src.operationLog || [],
              selectedNodeId: null, selectedEdgeId: null,
            });
            phaseData.acute    = loadPhase(obj.acute);
            phaseData.recovery = loadPhase(obj.recovery);
            // acuteRecord の復元
            phaseData.acuteRecord = {
              answers: {
                q4: obj.acuteRecord?.answers?.q4 || "",
                q5: obj.acuteRecord?.answers?.q5 || "",
              }
            };
            // recoveryCompare の復元（v4/v5 に存在、v3 は空で補完）
            phaseData.recoveryCompare = {
              answers: {
                q6:    obj.recoveryCompare?.answers?.q6    || "",
                q7:    obj.recoveryCompare?.answers?.q7    || "",
                q7sel: obj.recoveryCompare?.answers?.q7sel || "",
              }
            };
            // recoveryRecord の復元（v5 に存在、v3/v4 は空で補完）
            phaseData.recoveryRecord = {
              answers: {
                q8: obj.recoveryRecord?.answers?.q8 || "",
                q9: obj.recoveryRecord?.answers?.q9 || "",
              }
            };
            // phase5Data の復元（同一ラベルのノードを補完）
            if (obj.phase5Data?.removals) {
              const restoredRemovals = [];
              const seenLabels = new Set();
              for (const r of obj.phase5Data.removals) {
                if (seenLabels.has(r.label)) continue;
                seenLabels.add(r.label);
                const sameLabel = (window.idealMapAcute?.nodes || []).filter(n => n.label === r.label);
                if (sameLabel.length > 0) {
                  sameLabel.forEach(n => {
                    restoredRemovals.push({ nodeId: n.id, label: n.label, reason: r.reason || "" });
                  });
                } else {
                  // idealMapAcute未ロードの場合は元のエントリをそのまま保持
                  restoredRemovals.push(r);
                }
              }
              window.phase5Data = { removals: restoredRemovals };
            }
            // recovery → phaseData.p6 に復元（Phase6 の正規データ）
            phaseData.p6 = loadPhase(obj.recovery);
            phase6Initialized = phaseData.p6.nodes.length > 0;
            if (phase6Initialized) {
              // インポートした p6 データを有効とみなす。
              // この時点で phase5Data の復元も完了しているため、
              // getRemovalSignature() でシグネチャを記録しておくことで、
              // Phase6 入場時に p6NeedsRebuild()=false となり不要な再構築を防ぐ。
              phase6RemovalSignature = getRemovalSignature();
            }
            if (validateEdgeConflicts(phaseData.acute.edges) ||
                validateEdgeConflicts(phaseData.recovery.edges)) hasConflict = true;
          } else if (obj.version === 2 && obj.acute && obj.recovery) { // v2 後方互換
            const loadPhaseV2 = (src) => ({
              nodes: (src.nodes || []).map(n => ({ layerId: null, layerReason: "", isInitial: false, ...n })),
              edges: src.edges || [],
              answers: { q1: "", q2: "", p3q1: "", p3q2: "", p3q2sel: "", ...(src.answers || {}) },
              log: src.operationLog || [],
              selectedNodeId: null, selectedEdgeId: null,
            });
            phaseData.acute    = loadPhaseV2(obj.acute);
            phaseData.recovery = loadPhaseV2(obj.recovery);
            // v2 には acuteRecord / recoveryCompare / recoveryRecord がないため空で補完
            phaseData.acuteRecord     = { answers: { q4: "", q5: "" } };
            phaseData.recoveryCompare = { answers: { q6: "", q7: "", q7sel: "" } };
            phaseData.recoveryRecord  = { answers: { q8: "", q9: "" } };
            if (obj.phase5Data?.removals) {
              const restoredRemovals = [];
              const seenLabels = new Set();
              for (const r of obj.phase5Data.removals) {
                if (seenLabels.has(r.label)) continue;
                seenLabels.add(r.label);
                const sameLabel = (window.idealMapAcute?.nodes || []).filter(n => n.label === r.label);
                if (sameLabel.length > 0) {
                  sameLabel.forEach(n => {
                    restoredRemovals.push({ nodeId: n.id, label: n.label, reason: r.reason || "" });
                  });
                } else {
                  restoredRemovals.push(r);
                }
              }
              window.phase5Data = { removals: restoredRemovals };
            }
            phaseData.p6 = loadPhaseV2(obj.recovery);
            phase6Initialized = phaseData.p6.nodes.length > 0;
            if (phase6Initialized) {
              phase6RemovalSignature = getRemovalSignature();
            }
            if (validateEdgeConflicts(phaseData.acute.edges) ||
                validateEdgeConflicts(phaseData.recovery.edges)) hasConflict = true;
          } else if (Array.isArray(obj.nodes) && Array.isArray(obj.edges)) {
            // v1 legacy: 急性期に読み込む
            phaseData.acute = {
              nodes: obj.nodes.map(n => ({ layerId: null, layerReason: "", ...n })),
              edges: obj.edges,
              answers: { q1: "", q2: "", p3q1: "", p3q2: "", ...(obj.answers || {}) },
              log: obj.operationLog || [],
              selectedNodeId: null, selectedEdgeId: null,
            };
            // v1 には acuteRecord / recoveryCompare / recoveryRecord がないため空で補完
            phaseData.acuteRecord     = { answers: { q4: "", q5: "" } };
            phaseData.recoveryCompare = { answers: { q6: "", q7: "", q7sel: "" } };
            phaseData.recoveryRecord  = { answers: { q8: "", q9: "" } };
            if (validateEdgeConflicts(phaseData.acute.edges)) hasConflict = true;
          } else {
            throw new Error();
          }

          if (hasConflict) showToast("読み込んだデータに矛盾する矢印の組み合わせが含まれています", 3000);
          logOp("IMPORT", {});
          saveToLocalStorage();
          // 急性期フェーズに切り替えて表示
          // ※ switchPhase 内の savePhaseData が phaseData.recovery を上書きするため、
          //   インポート済みデータをスナップショットしてから復元する
          if (state.phase !== PHASE.ACUTE_MAP) {
            const _importedRecovery = phaseData.recovery;
            switchPhase(PHASE.ACUTE_MAP);
            phaseData.recovery = _importedRecovery;
          } else {
            loadPhaseData("acute");
            if ($("q2Answer")) $("q2Answer").value = state.answers.q2 || "";
            const cc = $("charCount"); if (cc) cc.textContent = (state.answers.q2 || "").length;
            renderAll();
          }
        } catch { alert("JSON形式が不正です。"); }
      };
      fr.readAsText(file);
    };
    inp.click();
  }

  function resetAll() {
    // 対応検証記録：問4・問5 の回答をクリア
    if (state.phase === PHASE.ACUTE_RECORD) {
      if (!confirm("問4・問5 の回答をリセットしますか？")) return;
      phaseData.acuteRecord.answers = { q4: "", q5: "" };
      renderAcuteRecordView();   // フォームを再描画してリセット状態に戻す
      return;
    }
    // 復旧期比較・分析：問6・問7 の回答をクリア
    if (state.phase === PHASE.RECOVERY_COMPARE) {
      if (!confirm("復旧期比較・分析の回答をリセットしますか？")) return;
      phaseData.recoveryCompare.answers = { q6: "", q7: "", q7sel: "" };
      const q6el = $("rcQ6Answer");
      const q7el = $("rcQ7Answer");
      if (q6el) { q6el.value = ""; const cc = $("rcQ6CharCount"); if (cc) cc.textContent = "0"; }
      if (q7el) { q7el.value = ""; const cc = $("rcQ7CharCount"); if (cc) cc.textContent = "0"; }
      document.querySelectorAll('input[name="rcQ7principle"]').forEach(r => { r.checked = false; });
      return;
    }
    // 復旧期対応検証記録：問8・問9 の回答をクリア
    if (state.phase === PHASE.RECOVERY_RECORD) {
      if (!confirm("問8・問9 の回答をリセットしますか？")) return;
      phaseData.recoveryRecord.answers = { q8: "", q9: "" };
      renderRecoveryRecordView();
      return;
    }
    // 復旧期準備：削除候補リストのみクリア
    if (state.phase === PHASE.RECOVERY_PREP) {
      if (!confirm("削除候補の選択をすべてリセットしますか？")) return;
      window.phase5Data.removals = [];
      // 削除候補がすべてなくなるため Phase6 も無効化する
      invalidatePhase6();
      renderPhase5Map();
      return;
    }
    // 復旧期マップ：ノード・矢印・ログのみクリア（再初期化を許可）
    if (state.phase === PHASE.RECOVERY_MAP) {
      if (!confirm("リセットしますか？（復旧期マップのノード・矢印・ログを消します）")) return;
      state.nodes = []; state.edges = [];
      state.answers = { q1: "", q2: "" }; state.log = [];
      state.selectedNodeId = null; state.selectedEdgeId = null;
      // Phase6 を無効化: 次回入場時に initPhase6Canvas() が再実行される
      invalidatePhase6();
      cancelArrowDraw();
      logOp("RESET", { phase: "p6" });
      renderAll();
      return;
    }
    const phaseLabel = state.phase === PHASE.RECOVERY_MAP ? "復旧期" : "急性期";
    if (!confirm(`初期化しますか？（${phaseLabel}のノード・矢印・回答・ログが消えます）`)) return;
    state.nodes = []; state.edges = [];
    state.answers = { q1: "", q2: "" }; state.log = [];
    state.selectedNodeId = null; state.selectedEdgeId = null;
    cancelArrowDraw();
    if ($("q2Answer")) $("q2Answer").value = "";
    const cc = $("charCount"); if (cc) cc.textContent = "0";
    logOp("RESET", { phase: activePhaseKey });
    renderAll();
  }

  // ================================================================
  // SCORING (reserved for future implementation)
  // ================================================================

  // ================================================================
  // READ-ONLY MAP RENDERER
  // ================================================================
  /**
   * 読み取り専用マップを指定キャンバスに描画する
   * @param {Array}       nodes        - 描画するノード配列
   * @param {Array}       edges        - 描画するエッジ配列
   * @param {HTMLElement} canvasEl     - ノードを配置するdiv
   * @param {SVGElement}  svgEl        - エッジを描画するSVG
   * @param {HTMLElement} canvasWrapEl - canvasWrap要素
   * @param {HTMLElement} statEl       - ノード数表示用span（nullも可）
   * @param {string}      markerSuffix - マーカーID用サフィックス（例: "-ideal"）
   */
  function renderReadOnlyMap(nodes, edges,
      canvasEl, svgEl, canvasWrapEl, statEl, markerSuffix,
      onNodeClick = null, fitToWrap = false,
      onNodeDblClick = null) {

    // 前回の fitToWrap transform が残っている場合に備えてリセット
    canvasEl.style.transform = "";
    canvasEl.style.transformOrigin = "";
    svgEl.style.transform = "";
    svgEl.style.transformOrigin = "";

    // ノード描画
    let roHighlightId = null;
    canvasEl.innerHTML = "";
    for (const n of nodes) {
      const div = document.createElement("div");
      const layerClass = n.layerId ? `layer-${n.layerId}` : "layer-none";
      // beneficiary 判定は label ベースに統一（renderNodes と同じ BENEFICIARY_LABELS を参照）
      // group クラスは class に含めない（内部メタデータ）
      const benefClass = BENEFICIARY_LABELS.has(n.label) ? " node-beneficiary" : "";
      div.className = `node ${layerClass}${benefClass}`;
      div.dataset.id = n.id;
      div.style.left = n.x + "px";
      div.style.top  = n.y + "px";
      div.innerHTML  = `<div class="ntitle">${esc(n.label)}</div>`;
      div.addEventListener("click", e => {
        e.stopPropagation();
        if (onNodeClick) onNodeClick(n.id, n.label);
      });
      div.addEventListener("dblclick", e => {
        e.stopPropagation();
        clearTimeout(_clickTimer);
        if (roHighlightId === n.id) {
          roHighlightId = null;
          clearHighlightRO(canvasEl, svgEl);
        } else {
          roHighlightId = n.id;
          onNodeDblClick?.();
          applyHighlightRO(n.id, nodes, edges, canvasEl, svgEl);
        }
      });
      canvasEl.appendChild(div);
    }

    // 背景クリックでハイライト解除
    canvasWrapEl.addEventListener("click", () => {
      if (roHighlightId !== null) {
        roHighlightId = null;
        clearHighlightRO(canvasEl, svgEl);
      }
    });

    // エッジ描画は1フレーム後（レイアウト確定後）に実行
    requestAnimationFrame(() => {
      const defs = svgEl.querySelector("defs");
      svgEl.innerHTML = "";
      if (defs) svgEl.appendChild(defs);

      for (const e of edges) {
        const from = nodes.find(n => n.id === e.from);
        const to   = nodes.find(n => n.id === e.to);
        if (!from || !to) continue;

        const getCenterRO = (node) => {
          const el = canvasEl.querySelector(`.node[data-id="${node.id}"]`);
          if (!el) return { x: node.x + 75, y: node.y + 28 };
          const r  = el.getBoundingClientRect();
          const cr = canvasWrapEl.getBoundingClientRect();
          return { x: r.left - cr.left + r.width/2, y: r.top - cr.top + r.height/2 };
        };

        const typeInfo = EDGE_MAP[e.label] || EDGE_TYPES[0];
        const col      = typeInfo.stroke;
        const a = getCenterRO(from);
        const b = getCenterRO(to);
        const dx = b.x - a.x, dy = b.y - a.y;
        const dist = Math.sqrt(dx*dx + dy*dy) || 1;
        const shorten = 26;
        const ex = b.x - (dx/dist)*shorten;
        const ey = b.y - (dy/dist)*shorten;

        const hasReverse = edges.some(r =>
          r.id !== e.id && r.from === e.to && r.to === e.from
        );

        let pathD, lx, ly;
        const CURVE_OFFSET = 50;

        if (hasReverse) {
          // renderEdges と同じく canonical 方向（ID小→大）を基準に法線を計算する。
          // エッジごとに dx/dy が反転するため、それをそのまま法線に使うと
          // sign の ±1 と打ち消し合い、両エッジが同じ側に湾曲して重なってしまう。
          const sign = e.from < e.to ? 1 : -1;
          const canonFrom = e.from < e.to ? from : to;
          const canonTo   = e.from < e.to ? to   : from;
          const ca = getCenterRO(canonFrom);
          const cb = getCenterRO(canonTo);
          const cdx = cb.x - ca.x, cdy = cb.y - ca.y;
          const cdist = Math.sqrt(cdx*cdx + cdy*cdy) || 1;
          const cnx = -cdy / cdist, cny = cdx / cdist;
          const sx = a.x + (dx/dist)*shorten;
          const sy = a.y + (dy/dist)*shorten;
          const cpx = (sx + ex) / 2 + sign * CURVE_OFFSET * cnx;
          const cpy = (sy + ey) / 2 + sign * CURVE_OFFSET * cny;
          pathD = `M ${sx} ${sy} Q ${cpx} ${cpy} ${ex} ${ey}`;
          lx = (sx + 2*cpx + ex)/4;
          ly = (sy + 2*cpy + ey)/4;
        } else if (e.bidirectional) {
          const sx = a.x + (dx/dist)*shorten;
          const sy = a.y + (dy/dist)*shorten;
          pathD = `M ${sx} ${sy} L ${ex} ${ey}`;
          lx = (a.x+b.x)/2; ly = (a.y+b.y)/2;
        } else {
          pathD = `M ${a.x} ${a.y} L ${ex} ${ey}`;
          lx = (a.x+b.x)/2; ly = (a.y+b.y)/2;
        }

        const markerKey = col === "#ff6b6b" ? "red"
                        : col === "#4d8fff" ? "blue"
                        : col === "#c084fc" ? "purple" : "teal";
        const markerEnd   = `url(#arrow-${markerKey}${markerSuffix})`;
        const markerStart = e.bidirectional ? markerEnd : "none";

        const g = document.createElementNS("http://www.w3.org/2000/svg","g");
        g.dataset.from = e.from; g.dataset.to = e.to;
        svgEl.appendChild(g);

        const path = document.createElementNS("http://www.w3.org/2000/svg","path");
        path.setAttribute("d", pathD);
        path.setAttribute("stroke", col);
        path.setAttribute("stroke-width", "2.2");
        path.setAttribute("fill", "none");
        path.setAttribute("marker-end", markerEnd);
        if (e.bidirectional) path.setAttribute("marker-start", markerStart);
        path.setAttribute("opacity", "0.88");
        g.appendChild(path);

        const labelW = (e.label?.length || 0) * 8 + 12;
        const bg = document.createElementNS("http://www.w3.org/2000/svg","rect");
        bg.setAttribute("x", lx-labelW/2); bg.setAttribute("y", ly-11);
        bg.setAttribute("width", labelW);  bg.setAttribute("height", 16);
        bg.setAttribute("rx", 5); bg.setAttribute("fill","#0d1422");
        bg.setAttribute("opacity","0.9");
        g.appendChild(bg);

        const txt = document.createElementNS("http://www.w3.org/2000/svg","text");
        txt.setAttribute("x", lx); txt.setAttribute("y", ly-2);
        txt.setAttribute("class","edge-label-text");
        txt.setAttribute("fill", col);
        txt.textContent = e.label || "";
        g.appendChild(txt);
      }

      // ノード数・矢印数の表示
      if (statEl) {
        statEl.textContent = `ノード: ${nodes.length} ／ 矢印: ${edges.length}`;
      }

      // フィット表示（Phase3のみ）
      if (fitToWrap && nodes.length > 0) {
        const xs = nodes.map(n => n.x);
        const ys = nodes.map(n => n.y);
        const minX = Math.min(...xs);
        const minY = Math.min(...ys);
        const maxX = Math.max(...xs) + 180;
        const maxY = Math.max(...ys) + 56;
        const contentW = maxX - minX;
        const contentH = maxY - minY;

        const wrapRect = canvasWrapEl.getBoundingClientRect();
        const scaleX = wrapRect.width  / contentW;
        const scaleY = wrapRect.height / contentH;
        const scale  = Math.min(scaleX, scaleY, 1) * 0.88;

        const tx = -minX * scale + (wrapRect.width  - contentW * scale) / 2;
        const ty = -minY * scale + (wrapRect.height - contentH * scale) / 2;

        const transform       = `translate(${tx}px, ${ty}px) scale(${scale})`;
        const transformOrigin = "top left";

        canvasEl.style.transform       = transform;
        canvasEl.style.transformOrigin = transformOrigin;
        svgEl.style.transform          = transform;
        svgEl.style.transformOrigin    = transformOrigin;
      }
    });
  }

  // ================================================================
  // PHASE 5 — READ-ONLY MAP & REMOVAL CANDIDATE LOGIC
  // ================================================================
  function renderPhase5Map() {
    if (mapLoadStatus.idealAcute !== "ready") {
      if (canvasEl) canvasEl.innerHTML = mapLoadStatus.idealAcute === "error"
        ? '<div style="color:var(--red);padding:20px;font-size:14px;">⚠ 理想マップの読み込みに失敗しました</div>'
        : '<div style="color:var(--text-dim);padding:20px;font-size:14px;">読み込み中…</div>';
      return;
    }
    state.nodes           = window.idealMapAcute.nodes.map(n => ({ ...n }));
    state.edges           = window.idealMapAcute.edges.map(e => ({ ...e }));
    state.selectedNodeId  = null;
    state.selectedEdgeId  = null;
    state.highlightNodeId = null;

    renderReadOnlyMap(
      state.nodes, state.edges,
      canvasEl, svgEl, canvasWrap,
      activeCanvasStatEl, activeMarkerSuffix,
      toggleRemovalCandidate
    );
    renderRemovalList();
    updatePhase6Btn();
  }

  function toggleRemovalCandidate(_nodeId, label) {
    // Phase6 が初期化済みの場合、削除候補の変更により phaseData.p6 が無効になる。
    // ・hasP6Edits() が true（ユーザーが編集済み）のときのみ confirm で確認する
    // ・confirm キャンセルならトグル自体を中断する
    // ・confirm 通過（または編集なし）なら invalidatePhase6() で Phase6 を無効化する
    if (phase6Initialized) {
      if (hasP6Edits()) {
        if (!confirm("削除候補を変更すると、復旧期マップの編集内容がリセットされます。続けますか？")) {
          return;
        }
      }
      // 削除候補が変わるため Phase6 を無効化する。
      // 次回 switchPhase(RECOVERY_MAP) → p6NeedsRebuild()=true で再構築される。
      invalidatePhase6();
    }

    const alreadySelected = window.phase5Data.removals.some(r => r.label === label);

    if (alreadySelected) {
      // 同一ラベルのエントリをすべて削除
      window.phase5Data.removals =
        window.phase5Data.removals.filter(r => r.label !== label);
      // 同一ラベルの全ノードからクラスを除去
      state.nodes
        .filter(n => n.label === label)
        .forEach(n => {
          canvasEl?.querySelector(`.node[data-id="${n.id}"]`)
            ?.classList.remove("node-removal-candidate");
        });
    } else {
      // 同一ラベルの全ノードをまとめて追加（重複防止あり）
      state.nodes
        .filter(n => n.label === label)
        .forEach(n => {
          if (!window.phase5Data.removals.some(r => r.nodeId === n.id)) {
            window.phase5Data.removals.push({ nodeId: n.id, label: n.label, reason: "" });
          }
          canvasEl?.querySelector(`.node[data-id="${n.id}"]`)
            ?.classList.add("node-removal-candidate");
        });
    }

    renderRemovalList();
    updatePhase6Btn();
  }

  function renderRemovalList() {
    const listEl = $("p5RemovalList");
    if (!listEl) return;
    const { removals } = window.phase5Data;
    if (removals.length === 0) {
      listEl.innerHTML = '<span style="color:var(--text-muted);font-size:12px;">（まだ選択されていません）</span>';
      return;
    }
    // ラベルで重複排除（先頭エントリを代表として使用）
    const uniqueRemovals = Object.values(
      removals.reduce((acc, r) => {
        if (!acc[r.label]) acc[r.label] = r;
        return acc;
      }, {})
    );
    listEl.innerHTML = uniqueRemovals.map((r) => `
      <div class="removal-item" data-label="${esc(r.label)}">
        <div class="removal-item-label">🗑 ${esc(r.label)}</div>
        <input type="text" class="removal-reason-input"
               maxlength="30" placeholder="なぜ不要か（任意・30字以内）"
               value="${esc(r.reason)}" data-label="${esc(r.label)}" />
      </div>
    `).join("");
    listEl.querySelectorAll(".removal-reason-input").forEach(inp => {
      inp.addEventListener("input", () => {
        // 同一ラベルの全エントリに同じ理由を反映
        const lbl = inp.dataset.label;
        window.phase5Data.removals
          .filter(r => r.label === lbl)
          .forEach(r => { r.reason = inp.value; });
      });
    });
  }

  function updatePhase6Btn() {
    const btn  = $("btnToPhase6");
    const hint = $("btnToPhase6Hint");
    if (!btn) return;
    const has = window.phase5Data.removals.length > 0;
    btn.disabled = !has;
    if (hint) hint.style.display = has ? "none" : "block";
  }

  window.goToPhase6 = function() { switchPhase(PHASE.RECOVERY_MAP); };

  // ================================================================
  // PHASE 6 — CANVAS INITIALISATION
  // ================================================================
  function initPhase6Canvas() {
    if (mapLoadStatus.idealAcute !== "ready") return;
    const removedIds = new Set(window.phase5Data.removals.map(r => r.nodeId));

    // 旧ID → 新ID のマッピング（canvas-p5 との data-id 重複を解消）
    const idMap = new Map();
    window.idealMapAcute.nodes
      .filter(n => !removedIds.has(n.id))
      .forEach(n => idMap.set(n.id, "n-" + uid()));

    phaseData.p6.nodes = window.idealMapAcute.nodes
      .filter(n => !removedIds.has(n.id))
      .map(n => ({
        id:          idMap.get(n.id),
        label:       n.label,
        group:       n.group,
        x:           n.x,
        y:           n.y,
        layerId:     PHASE6_BENEFICIARY_LABELS.has(n.label) ? 4 : null,
        layerReason: "",
        isInitial:   true,
      }));

    phaseData.p6.edges = [];

    phaseData.p6.answers = { q1: "", q2: "" };
    phaseData.p6.log     = [];
    // initPhase6Canvas() は Phase6 の「初期状態を作る」唯一の責務を持つ。
    // 呼び出し後は phase6Initialized=true + シグネチャが揃い、p6NeedsRebuild()=false になる。
    phase6Initialized      = true;
    phase6RemovalSignature = getRemovalSignature();
    logOp("INIT_PHASE6", { nodeCount: phaseData.p6.nodes.length });
  }

  // ================================================================
  // ================================================================
  // ICS原則 選択表示ヘルパー
  // ================================================================
  const ICS_PRINCIPLE_JA = {
    "Unity of Command":     "指揮一元化",
    "Unified Command":      "統合指揮",
    "Span of Control":      "統制範囲",
    "Modular Organization": "組立型組織",
    "Communications":       "コミュニケーション",
  };

  function renderSelectedPrinciple(elementId, value, fallbackMsg) {
    const el = $(elementId);
    if (!el) return;
    if (!value) {
      el.innerHTML = `<span style="font-size:11px;color:var(--text-muted);">${esc(fallbackMsg || "（未選択）")}</span>`;
      return;
    }
    const ja = ICS_PRINCIPLE_JA[value] || "";
    el.innerHTML = `
      <div class="ar-principle-chip">
        <span class="ar-principle-en">${esc(value)}</span>
        <span class="ar-principle-ja">${esc(ja)}</span>
      </div>
    `;
  }

  // ACUTE_RECORD フェーズ レンダラー [ADDED]
  // ================================================================

  /**
   * 問4 用ラジオ選択肢 HTML を返す（excerpts 配列から自動生成）
   * @param {Object} q - ACUTE_RECORD_CONTENT.questions の問オブジェクト
   * @param {Array}  excerpts - ACUTE_RECORD_CONTENT.excerpts
   * @returns {string} HTML 文字列
   */
  function renderSingleChoiceQuestion(q, excerpts) {
    const options = excerpts.map(ex => `
      <label class="ics-radio-item">
        <input type="radio" name="arQ4" value="${esc(ex.id)}">
        <span class="ics-radio-label">${esc(ex.id)}</span>
      </label>
    `).join("");
    return `
      <div class="ar-question-block" id="arQBlock-${esc(q.id)}">
        <div class="compare-qa-label">${esc(q.label)}</div>
        <div class="ics-radio-group" id="arQ4RadioGroup">
          ${options}
        </div>
      </div>
    `;
  }

  /**
   * 問5 用テキストエリア HTML を返す
   * @param {Object} q - ACUTE_RECORD_CONTENT.questions の問オブジェクト
   * @returns {string} HTML 文字列
   */
  function renderTextareaQuestion(q) {
    // FIXME: textarea 問題が将来複数になる場合は question.id ベースで ID を動的生成すること。
    //        現状は1問固定前提でハードコードしている。
    return `
      <div class="ar-question-block" id="arQBlock-${esc(q.id)}">
        <div class="compare-qa-label">${esc(q.label)}</div>
        <textarea id="arQ5Answer" rows="3"
          placeholder="${esc(q.placeholder || "")}"
          maxlength="${q.maxLength || 500}"></textarea>
        <div class="char-count"><span id="arQ5CharCount">0</span> / ${q.maxLength || 500} 字</div>
      </div>
    `;
  }

  /**
   * ACUTE_RECORD フェーズの画面を描画し、フォームイベントを attach する。
   * wireEvents() には追加しない（DOM が動的生成のため）。
   */
  function renderAcuteRecordView() {
    const excerptList = $("arExcerptList");
    const questionArea = $("arQuestionArea");
    if (!excerptList || !questionArea) return;

    const { excerpts, questions } = ACUTE_RECORD_CONTENT;

    // (a) 抜粋エリア描画
    excerptList.innerHTML = excerpts.map(ex => `
      <div class="ar-excerpt-card" data-id="${esc(ex.id)}" id="arCard-${esc(ex.id)}">
        <span class="ar-excerpt-num">${esc(ex.id)}</span>
        <div class="ar-excerpt-text">${esc(ex.text)}</div>
      </div>
    `).join("");

    // (b) 設問エリア描画
    questionArea.innerHTML = questions.map(q => {
      if (q.kind === "singleChoice") return renderSingleChoiceQuestion(q, excerpts);
      if (q.kind === "textarea")     return renderTextareaQuestion(q);
      return "";
    }).join("");

    // (c) イベント attach — q4 ラジオ
    questionArea.querySelectorAll('input[name="arQ4"]').forEach(radio => {
      radio.addEventListener("change", () => {
        phaseData.acuteRecord.answers.q4 = radio.value;
        // 抜粋カードのハイライト連動
        excerptList.querySelectorAll(".ar-excerpt-card").forEach(card => {
          card.classList.toggle("selected", card.dataset.id === radio.value);
        });
        saveToLocalStorage();
      });
    });

    // (c) イベント attach — q5 テキストエリア（文字数カウント付き）
    const q5ta = $("arQ5Answer");
    if (q5ta) {
      q5ta.addEventListener("input", () => {
        const len = q5ta.value.length;
        const cc  = $("arQ5CharCount");
        if (cc) {
          cc.textContent = len;
          const max = ACUTE_RECORD_CONTENT.questions.find(q => q.id === "q5")?.maxLength || 200;
          cc.parentElement.className =
            "char-count" + (len > max ? " over" : len >= max * 0.9 ? " warn" : "");
          if (len > max) {
            q5ta.classList.add("over");
          } else {
            q5ta.classList.remove("over");
          }
        }
        phaseData.acuteRecord.answers.q5 = q5ta.value;
        debouncedSave();
      });
    }

    // 「復旧期準備へ進む」ボタン（バリデーション付き）
    const btnToRecoveryPrep = $("btnToRecoveryPrep");
    if (btnToRecoveryPrep) {
      // 重複イベント防止のためクローン置き換え
      const fresh = btnToRecoveryPrep.cloneNode(true);
      btnToRecoveryPrep.replaceWith(fresh);
      fresh.addEventListener("click", () => {
        const { q4, q5 } = phaseData.acuteRecord.answers;
        if (!q4) {
          showToast("問4で対応検証記録の番号を選択してください", 3000);
          return;
        }
        if (!q5) {
          if (!confirm("問5が未入力です。このまま進みますか？")) return;
        }
        switchPhase(PHASE.RECOVERY_PREP);
      });
    }
  }

  /**
   * phaseData.acuteRecord.answers の値を各フォーム要素に復元する。
   * renderAcuteRecordView() の DOM 生成後に呼ぶこと。
   */
  function restoreAcuteRecordAnswers() {
    const { q4, q5 } = phaseData.acuteRecord.answers;
    const excerptList = $("arExcerptList");

    // q4 復元
    if (q4) {
      const radio = document.querySelector(`input[name="arQ4"][value="${q4}"]`);
      if (radio) {
        radio.checked = true;
        // カードハイライト
        excerptList?.querySelectorAll(".ar-excerpt-card").forEach(card => {
          card.classList.toggle("selected", card.dataset.id === q4);
        });
      }
    }

    // q5 復元
    const q5ta = $("arQ5Answer");
    if (q5ta && q5) {
      q5ta.value = q5;
      const cc = $("arQ5CharCount");
      if (cc) {
        cc.textContent = q5.length;
        const max = ACUTE_RECORD_CONTENT.questions.find(q => q.id === "q5")?.maxLength || 200;
        cc.parentElement.className =
          "char-count" + (q5.length > max ? " over" : q5.length >= max * 0.9 ? " warn" : "");
      }
    }
  }

  // ================================================================
  // RECOVERY_RECORD フェーズ レンダラー
  // ================================================================

  function renderRecoveryRecordView() {
    const excerptList = $("rrExcerptList");
    const questionArea = $("rrQuestionArea");
    if (!excerptList || !questionArea) return;

    const { excerpts, questions } = RECOVERY_RECORD_CONTENT;

    // (a) 抜粋エリア描画
    excerptList.innerHTML = excerpts.map(ex => `
      <div class="ar-excerpt-card" data-id="${esc(ex.id)}" id="rrCard-${esc(ex.id)}">
        <span class="ar-excerpt-num">${esc(ex.id)}</span>
        <div class="ar-excerpt-text">${esc(ex.text)}</div>
      </div>
    `).join("");

    // (b) 設問エリア描画
    questionArea.innerHTML = questions.map(q => {
      if (q.kind === "singleChoice") {
        const options = excerpts.map(ex => `
          <label class="ics-radio-item">
            <input type="radio" name="rrQ8" value="${esc(ex.id)}">
            <span class="ics-radio-label">${esc(ex.id)}</span>
          </label>
        `).join("");
        return `
          <div class="ar-question-block" id="rrQBlock-${esc(q.id)}">
            <div class="compare-qa-label">${esc(q.label)}</div>
            <div class="ics-radio-group" id="rrQ8RadioGroup">
              ${options}
            </div>
          </div>
        `;
      }
      if (q.kind === "textarea") {
        // FIXME: textarea 問題が将来複数になる場合は question.id ベースで ID を動的生成すること。
        //        現状は1問固定前提でハードコードしている。
        return `
          <div class="ar-question-block" id="rrQBlock-${esc(q.id)}">
            <div class="compare-qa-label">${esc(q.label)}</div>
            <textarea id="rrQ9Answer" rows="3"
              placeholder="${esc(q.placeholder || "")}"
              maxlength="${q.maxLength || 500}"></textarea>
            <div class="char-count"><span id="rrQ9CharCount">0</span> / ${q.maxLength || 500} 字</div>
          </div>
        `;
      }
      return "";
    }).join("");

    // (c) イベント attach — q8 ラジオ
    questionArea.querySelectorAll('input[name="rrQ8"]').forEach(radio => {
      radio.addEventListener("change", () => {
        phaseData.recoveryRecord.answers.q8 = radio.value;
        excerptList.querySelectorAll(".ar-excerpt-card").forEach(card => {
          card.classList.toggle("selected", card.dataset.id === radio.value);
        });
        saveToLocalStorage();
      });
    });

    // (c) イベント attach — q9 テキストエリア
    const q9ta = $("rrQ9Answer");
    if (q9ta) {
      q9ta.addEventListener("input", () => {
        const len = q9ta.value.length;
        const cc  = $("rrQ9CharCount");
        if (cc) {
          cc.textContent = len;
          const max = RECOVERY_RECORD_CONTENT.questions.find(q => q.id === "q9")?.maxLength || 200;
          cc.parentElement.className =
            "char-count" + (len > max ? " over" : len >= max * 0.9 ? " warn" : "");
          q9ta.classList.toggle("over", len > max);
        }
        phaseData.recoveryRecord.answers.q9 = q9ta.value;
        debouncedSave();
      });
    }

    // 「シーケンス図へ進む」ボタン
    const btnToSeq = $("btnToSequenceFromRR");
    if (btnToSeq) {
      const fresh = btnToSeq.cloneNode(true);
      btnToSeq.replaceWith(fresh);
      fresh.addEventListener("click", () => {
        const { q8, q9 } = phaseData.recoveryRecord.answers;
        if (!q8) {
          showToast("問8で対応検証記録の番号を選択してください", 3000);
          return;
        }
        if (!q9) {
          if (!confirm("問9が未入力です。このまま進みますか？")) return;
        }
        switchPhase(PHASE.SEQUENCE);
      });
    }
  }

  function restoreRecoveryRecordAnswers() {
    const { q8, q9 } = phaseData.recoveryRecord.answers;
    const excerptList = $("rrExcerptList");

    // q8 復元
    if (q8) {
      const radio = document.querySelector(`input[name="rrQ8"][value="${q8}"]`);
      if (radio) {
        radio.checked = true;
        excerptList?.querySelectorAll(".ar-excerpt-card").forEach(card => {
          card.classList.toggle("selected", card.dataset.id === q8);
        });
      }
    }

    // q9 復元
    const q9ta = $("rrQ9Answer");
    if (q9ta && q9) {
      q9ta.value = q9;
      const cc = $("rrQ9CharCount");
      if (cc) {
        cc.textContent = q9.length;
        const max = RECOVERY_RECORD_CONTENT.questions.find(q => q.id === "q9")?.maxLength || 200;
        cc.parentElement.className =
          "char-count" + (q9.length > max ? " over" : q9.length >= max * 0.9 ? " warn" : "");
      }
    }
  }

  // ================================================================
  // RENDER ALL
  // ================================================================
  function renderAll() {
    updateSelectors();
    renderNodes();
    renderEdges();
    updateCanvasStat();
    if (activePhaseKey === "p6") renderPalette();
    if (state.highlightNodeId) applyHighlight(state.highlightNodeId);
  }

  // ================================================================
  // EVENTS
  // ================================================================
  function wireEvents() {
    $("btnExport")?.addEventListener("click", exportJSON);
    $("btnImport")?.addEventListener("click", importJSON);
    $("btnReset")?.addEventListener("click", resetAll);

    // [ADDED] 「問4・5へ進む」ボタン（ACUTE_COMPARE → ACUTE_RECORD）
    $("btnToAcuteRecord")?.addEventListener("click", () => {
      const ans = phaseData.acute.answers;
      const allFilled = ans.p3q1 && ans.p3q2sel && ans.p3q2;
      if (!allFilled) {
        if (!confirm("未入力の設問があります。続けますか？")) return;
      }
      switchPhase(PHASE.ACUTE_RECORD);
    });

    // Phase3 記述問題パネル
    const p3q1 = $("p3q1Answer");
    const p3q2 = $("p3q2Answer");

    if (p3q1) {
      p3q1.addEventListener("input", () => {
        const len = p3q1.value.length;
        const cc  = $("p3q1CharCount");
        if (cc) {
          cc.textContent = len;
          cc.parentElement.className =
            "char-count" + (len > 100 ? " over" : len >= 90 ? " warn" : "");
        }
        phaseData.acute.answers.p3q1 = p3q1.value;
        debouncedSave();
      });
    }

    if (p3q2) {
      p3q2.addEventListener("input", () => {
        const len = p3q2.value.length;
        const cc  = $("p3q2CharCount");
        if (cc) {
          cc.textContent = len;
          cc.parentElement.className =
            "char-count" + (len > 100 ? " over" : len >= 90 ? " warn" : "");
        }
        phaseData.acute.answers.p3q2 = p3q2.value;
        debouncedSave();
      });
    }

    document.querySelectorAll('input[name="p3q2principle"]').forEach(radio => {
      radio.addEventListener("change", () => {
        phaseData.acute.answers.p3q2sel = radio.value;
      });
    });

    // [ADDED] 復旧期比較・分析 テキストエリア文字数カウント
    const rcQ6 = $("rcQ6Answer");
    const rcQ7 = $("rcQ7Answer");
    if (rcQ6) {
      rcQ6.addEventListener("input", () => {
        const len = rcQ6.value.length;
        const cc  = $("rcQ6CharCount");
        if (cc) {
          cc.textContent = len;
          const max = RECOVERY_COMPARE_CONTENT.questions.find(q => q.id === "q6")?.maxLength || 100;
          cc.parentElement.className = "char-count" + (len > max ? " over" : len >= max * 0.9 ? " warn" : "");
        }
        phaseData.recoveryCompare.answers.q6 = rcQ6.value;
        debouncedSave();
      });
    }
    if (rcQ7) {
      rcQ7.addEventListener("input", () => {
        const len = rcQ7.value.length;
        const cc  = $("rcQ7CharCount");
        if (cc) {
          cc.textContent = len;
          const max = RECOVERY_COMPARE_CONTENT.questions.find(q => q.id === "q7")?.maxLength || 200;
          cc.parentElement.className = "char-count" + (len > max ? " over" : len >= max * 0.9 ? " warn" : "");
        }
        phaseData.recoveryCompare.answers.q7 = rcQ7.value;
        debouncedSave();
      });
    }
    document.querySelectorAll('input[name="rcQ7principle"]').forEach(radio => {
      radio.addEventListener("change", () => {
        phaseData.recoveryCompare.answers.q7sel = radio.value;
      });
    });

    // 「対応検証記録（復旧期）へ進む」ボタン（問6 入力チェック付き）
    $("btnToRecoveryRecord")?.addEventListener("click", () => {
      const { q6, q7 } = phaseData.recoveryCompare.answers;
      if (!q6) {
        showToast("問6で構造的差異を入力してください", 3000);
        return;
      }
      if (!q7) {
        if (!confirm("問7 が未入力です。このまま進みますか？")) return;
      }
      switchPhase(PHASE.RECOVERY_RECORD);
    });

    document.addEventListener("keydown", e => {
      if (e.key === "Escape" && !state.drawingArrow && state.selectedNodeId) {
        clearSelection();
        return;
      }
      if (!state.selectedEdgeId) return;
      if (e.key !== "Delete" && e.key !== "Backspace") return;
      const tag = document.activeElement.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return;
      e.preventDefault();
      deleteEdge(state.selectedEdgeId);
    });
  }

  // ================================================================
  // TOAST NOTIFICATION
  // ================================================================
  function showToast(message, duration = 2000) {
    let toast = document.getElementById("toastNotification");
    if (!toast) {
      toast = document.createElement("div");
      toast.id = "toastNotification";
      Object.assign(toast.style, {
        position: "fixed", bottom: "80px", left: "50%",
        transform: "translateX(-50%)",
        background: "#1e1028", border: "1px solid #c084fc", color: "#d8b4fe",
        padding: "10px 22px", borderRadius: "10px", fontSize: "13px",
        fontWeight: "600", zIndex: "9999", whiteSpace: "nowrap",
        boxShadow: "0 4px 20px rgba(0,0,0,0.6)", transition: "opacity 0.3s",
        pointerEvents: "none",
      });
      document.body.appendChild(toast);
    }
    toast.textContent = message;
    toast.style.opacity = "1";
    clearTimeout(toast._timer);
    toast._timer = setTimeout(() => { toast.style.opacity = "0"; }, duration);
  }

  // ================================================================
  // EDGE EXCLUSIVITY CHECK
  // ================================================================
  /**
   * 2ノード間に指定タイプのエッジを追加できるか判定する
   * @param {string} fromId - 起点ノードID
   * @param {string} toId   - 終点ノードID
   * @param {string} type   - 追加しようとするエッジ種別 ("指示命令"|"情報伝達"|"連携協力"|"支援")
   * @param {Array}  edges  - 現在のエッジ配列
   * @returns {{ allowed: boolean, reason: string }}
   */
  function canAddEdge(fromId, toId, type, edges) {
    const CMD_INFO = new Set(["指示命令", "情報伝達"]);
    const COOP     = "連携協力";

    // 支援エッジは排他チェック対象外
    if (type === "支援") return { allowed: true, reason: "" };

    // 無方向ペアのエッジ（COOP共存チェック用）
    const pairEdges = edges.filter(e =>
      (e.from === fromId && e.to === toId) ||
      (e.from === toId   && e.to === fromId)
    );

    // COOP追加時
    if (type === COOP) {
      // CMD/INFOが存在すれば拒否（無方向）
      if (pairEdges.some(e => CMD_INFO.has(e.label))) {
        return { allowed: false, reason: "指示命令または情報伝達が設定済みのペアには連携協力を追加できません" };
      }
      // COOP重複チェックは無方向（A↔B を同一ペアとみなす）
      if (pairEdges.some(e => e.label === COOP)) {
        return { allowed: false, reason: "同じ種類の矢印がすでに存在します" };
      }
    }

    // CMD/INFO追加時
    if (CMD_INFO.has(type)) {
      // COOPが存在すれば拒否（無方向）
      if (pairEdges.some(e => e.label === COOP)) {
        return { allowed: false, reason: "連携協力が設定済みのペアには指示命令・情報伝達を追加できません" };
      }
      // 重複チェックは有方向（同一 from/to/type のみブロック。逆向きは別エッジとして許可）
      if (edges.some(e => e.from === fromId && e.to === toId && e.label === type)) {
        return { allowed: false, reason: "同じ種類の矢印がすでに存在します" };
      }
    }

    return { allowed: true, reason: "" };
  }

  // ================================================================
  // MAP LOADERS
  // ================================================================
  async function loadIdealMapAcute() {
    mapLoadStatus.idealAcute = "loading";
    try {
      const resp = await fetch('./ideal_map_acute.json');
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      window.idealMapAcute = await resp.json();
      mapLoadStatus.idealAcute = "ready";
    } catch (e) {
      console.error('[ICS] ideal_map_acute.json の読み込みに失敗:', e);
      if (window.IDEAL_MAP_ACUTE_FALLBACK) {
        window.idealMapAcute = window.IDEAL_MAP_ACUTE_FALLBACK;
        mapLoadStatus.idealAcute = "ready";
      } else {
        mapLoadStatus.idealAcute = "error";
      }
    }
    // Phase5 / Phase6 に滞在中なら自動再描画
    if (state.phase === PHASE.RECOVERY_PREP || state.phase === PHASE.RECOVERY_MAP) {
      switchPhase(state.phase);
    }
  }

  async function loadActualMapAcute() {
    mapLoadStatus.actualAcute = "loading";
    try {
      const resp = await fetch('./actual_map_acute.json');
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      window.actualMapAcute = await resp.json();
      mapLoadStatus.actualAcute = "ready";
    } catch (e) {
      console.error('[ICS] actual_map_acute.json の読み込みに失敗:', e);
      if (window.ACTUAL_MAP_ACUTE_FALLBACK) {
        window.actualMapAcute = window.ACTUAL_MAP_ACUTE_FALLBACK;
        mapLoadStatus.actualAcute = "ready";
      } else {
        mapLoadStatus.actualAcute = "error";
      }
    }
    // ACUTE_COMPARE に滞在中なら自動再描画
    if (state.phase === PHASE.ACUTE_COMPARE) {
      switchPhase(state.phase);
    }
  }

  async function loadActualMapRecovery() {
    mapLoadStatus.actualRecovery = "loading";
    try {
      const resp = await fetch('./actual_map_recovery.json');
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      window.actualMapRecovery = await resp.json();
      mapLoadStatus.actualRecovery = "ready";
    } catch (e) {
      console.error('[ICS] actual_map_recovery.json の読み込みに失敗:', e);
      mapLoadStatus.actualRecovery = "error";
    }
    // RECOVERY_COMPARE に滞在中なら自動再描画
    if (state.phase === PHASE.RECOVERY_COMPARE) {
      switchPhase(state.phase);
    }
  }

  // ================================================================
  // INIT
  // ================================================================
  function renderOrientationNodeList() {
    const container = $("orientationNodeList");
    if (!container) return;
    const ORIENT_BENEFICIARY = new Set(["避難所", "医療機関", "福祉避難所", "在宅避難者", "仮設住宅"]);
    const iconMap = new Map();
    for (const n of PALETTE_NODES)          { if (!iconMap.has(n.label)) iconMap.set(n.label, n.icon); }
    for (const n of RECOVERY_PALETTE_NODES) { if (!iconMap.has(n.label)) iconMap.set(n.label, n.icon); }
    const seen = new Set();
    const regular = [], beneficiary = [];
    for (const n of [...PALETTE_NODES, ...RECOVERY_PALETTE_NODES]) {
      if (seen.has(n.label)) continue;
      seen.add(n.label);
      (ORIENT_BENEFICIARY.has(n.label) ? beneficiary : regular).push(n.label);
    }
    const makeChip = (label, isBenef) => {
      const icon = iconMap.get(label) || "";
      const desc = NODE_DESCRIPTIONS[label] || "";
      const chip = document.createElement("div");
      chip.className = "member-chip" + (isBenef ? " chip-beneficiary" : "");
      chip.setAttribute("aria-label", `${label}：${desc}`);
      chip.innerHTML = `<div class="mchip-main"><span class="ico">${icon}</span>${esc(label)}</div><span class="mchip-desc">${esc(desc)}</span>`;
      return chip;
    };
    container.innerHTML = "";
    for (const label of regular) container.appendChild(makeChip(label, false));
    const sep = document.createElement("div");
    sep.className = "mchip-separator";
    sep.textContent = "支援対象";
    container.appendChild(sep);
    for (const label of beneficiary) container.appendChild(makeChip(label, true));
  }

  function init() {
    const cfg = MAP_PHASE_CONFIG[1];  // 急性期
    activePhaseKey     = cfg.key;
    activePaletteNodes = cfg.paletteNodes;
    BENEFICIARY_LABELS = cfg.beneficiaries;
    canvasEl           = $(cfg.domIds.canvas);
    svgEl              = $(cfg.domIds.svg);
    paletteEl          = $(cfg.domIds.palette);
    canvasWrap         = $(cfg.domIds.wrap);
    activeCanvasStatEl = $(cfg.domIds.stat);
    activeArrowHintEl  = $(cfg.domIds.hint);
    activeMarkerSuffix = cfg.markerSuffix;

    renderPalette();
    wireEvents();
    renderOrientationNodeList();
    renderAll();
    logOp("INIT", { scenarioId: SCENARIO.id });
    loadIdealMapAcute();
    loadActualMapAcute();
    loadActualMapRecovery(); // [ADDED]
  }

  init();

  // ── localStorage 復元フロー ────────────────────────────────────────
  (function restoreFromStorage() {
    const saved = loadFromLocalStorage();
    if (!saved) return;
    const savedAt = new Date(saved.savedAt).toLocaleString();
    if (confirm(`前回の作業（${savedAt} 保存）を復元しますか？\n「キャンセル」を選ぶと前回の作業は破棄され、新しいセッションとして開始します。`)) {
      Object.assign(phaseData, saved.phaseData);
      if (saved.phase5Data) Object.assign(window.phase5Data, saved.phase5Data);
      switchPhase(saved.currentPhase ?? PHASE.ORIENTATION);
    } else {
      clearLocalStorage();
    }
  })();

  // 被験者切替・実験者向け運用 API
  window.__icsClearStorage = clearLocalStorage;

  window.addEventListener("beforeunload", (e) => {
    if (hasUnsavedWork()) {
      e.preventDefault();
      e.returnValue = ""; // Chrome が要求
    }
  });
})();
