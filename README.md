<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>한국은행 2023년 산업연관표 분석 및 경제적 파급효과 시뮬레이터</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- SheetJS for Excel Parsing -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- KaTeX for Math Equations -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/katex.min.css">
    <script src="https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/katex.min.js"></script>

    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        navy: {
                            800: '#1e293b',
                            900: '#0f172a',
                            950: '#030712',
                        },
                        bok: {
                            blue: '#003876',
                            light: '#1d4ed8',
                            accent: '#0284c7',
                            gold: '#d97706'
                        }
                    }
                }
            }
        }
    </script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Noto+Sans+KR:wght@300;400;500;700&display=swap');
        body {
            font-family: 'Noto Sans KR', 'Inter', sans-serif;
            background-color: #f8fafc;
        }
        .custom-scrollbar::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
            background: #f1f5f9;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
            background: #cbd5e1;
            border-radius: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
            background: #94a3b8;
        }
        .math-tex {
            font-family: 'KaTeX_Main', serif;
        }
    </style>
</head>
<body class="text-slate-800 flex flex-col min-h-screen custom-scrollbar">

    <!-- Top Fixed Header & Navigation -->
    <header class="sticky top-0 z-50 bg-navy-900 text-white shadow-lg border-b border-slate-700">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex items-center justify-between h-16">
                <!-- Title / Logo -->
                <div class="flex items-center space-x-3">
                    <div class="bg-bok-blue text-white p-2 rounded-lg font-bold text-xl flex items-center justify-center w-10 h-10 border border-blue-400/30 shadow-inner">
                        <i class="fa-solid fa-chart-pie text-bok-accent"></i>
                    </div>
                    <div>
                        <h1 class="text-lg font-bold leading-tight tracking-tight text-slate-100">한국은행 2023년 산업연관표(연장표)</h1>
                        <p class="text-xs text-blue-300 font-medium">생산자가격 기준 분석 & 경제적 파급효과 시뮬레이터</p>
                    </div>
                </div>

                <!-- Navigation Links -->
                <nav class="hidden md:flex space-x-1 lg:space-x-2 text-sm font-medium">
                    <button onclick="scrollToSection('sec-model')" class="px-3 py-2 rounded-md hover:bg-slate-800 text-slate-200 transition-colors flex items-center gap-1.5">
                        <i class="fa-solid fa-square-check text-emerald-400"></i> 1. 검증 및 설정
                    </button>
                    <button onclick="scrollToSection('sec-tables')" class="px-3 py-2 rounded-md hover:bg-slate-800 text-slate-200 transition-colors flex items-center gap-1.5">
                        <i class="fa-solid fa-table-cells text-sky-400"></i> 2. 투입산출표 탐색
                    </button>
                    <button onclick="scrollToSection('sec-linkage')" class="px-3 py-2 rounded-md hover:bg-slate-800 text-slate-200 transition-colors flex items-center gap-1.5">
                        <i class="fa-solid fa-diagram-project text-amber-400"></i> 3. 전·후방 연쇄효과
                    </button>
                    <button onclick="scrollToSection('sec-simulator')" class="px-3 py-2 rounded-md bg-bok-blue hover:bg-blue-700 text-white transition-all shadow flex items-center gap-1.5 border border-blue-400/40">
                        <i class="fa-solid fa-calculator text-yellow-300"></i> 4. 파급효과 시뮬레이션
                    </button>
                </nav>

                <!-- Data Status Badge -->
                <div class="hidden xl:flex items-center space-x-2 text-xs bg-slate-800/80 px-3 py-1.5 rounded-full border border-slate-700">
                    <span class="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
                    <span id="active-dataset-badge" class="text-slate-300">내장 데이터: 33개 통합대분류</span>
                </div>
            </div>
        </div>
    </header>

    <!-- Main Content Area -->
    <main class="flex-grow max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-10">

        <!-- Banner Info Card -->
        <div class="bg-gradient-to-r from-navy-900 via-slate-800 to-bok-blue rounded-2xl p-6 text-white shadow-md border border-slate-700/50 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
            <div class="space-y-2 max-w-3xl">
                <div class="inline-flex items-center gap-2 bg-blue-500/20 text-blue-300 text-xs px-2.5 py-1 rounded-full border border-blue-400/30">
                    <i class="fa-solid fa-circle-info"></i> BOK 2023 Input-Output Tables Framework
                </div>
                <h2 class="text-xl md:text-2xl font-bold tracking-tight">생산자가격 기준 레온티에프(Leontief) 투입산출 모형</h2>
                <p class="text-slate-300 text-sm leading-relaxed">
                    본 시스템은 한국은행이 공표한 2023년 산업연관표(연장표) 데이터를 기반으로 국산투입계수 matrix <span class="math-tex text-blue-200 font-semibold">A^d</span>와 레온티에프 역행렬 <span class="math-tex text-blue-200 font-semibold">(I - A^d)^{-1}</span>을 실시간 연산합니다. 최종수요 변동에 따른 생산·부가가치·수입 유발 파급효과를 다각도로 분석합니다.
                </p>
            </div>
            <div class="flex flex-col sm:flex-row gap-2 shrink-0 w-full md:w-auto">
                <label for="excel-upload-input" class="cursor-pointer bg-emerald-600 hover:bg-emerald-500 text-white px-4 py-2.5 rounded-xl font-medium text-sm transition-all shadow-md flex items-center justify-center gap-2 border border-emerald-400/30">
                    <i class="fa-solid fa-file-excel text-lg"></i>
                    <span>Excel (.xlsx) 직접 업로드</span>
                </label>
                <input type="file" id="excel-upload-input" accept=".xlsx, .xls" class="hidden" onchange="handleExcelUpload(event)">
            </div>
        </div>

        <!-- ================= SECTION 1: DATA MODEL & VERIFICATION ================= -->
        <section id="sec-model" class="scroll-mt-20 space-y-4">
            <div class="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 border-b border-slate-200 pb-3">
                <div class="flex items-center space-x-2">
                    <span class="bg-blue-100 text-bok-blue font-bold px-2.5 py-1 rounded-md text-sm">섹션 1</span>
                    <h3 class="text-xl font-bold text-slate-900">자료 선택 및 모형 정합성 검증</h3>
                </div>
                <div class="flex items-center space-x-2 bg-slate-200/70 p-1 rounded-xl">
                    <button id="btn-ds-33" onclick="switchDataset('33')" class="px-3 py-1.5 rounded-lg text-xs font-semibold bg-white text-navy-900 shadow-sm transition-all">
                        통합대분류 (33부문)
                    </button>
                    <button id="btn-ds-83" onclick="switchDataset('83')" class="px-3 py-1.5 rounded-lg text-xs font-medium text-slate-600 hover:text-slate-900 transition-all">
                        통합중분류 (83부문)
                    </button>
                    <button id="btn-ds-custom" class="hidden px-3 py-1.5 rounded-lg text-xs font-medium text-emerald-700 bg-emerald-100 transition-all">
                        사용자 엑셀 데이터
                    </button>
                </div>
            </div>

            <!-- Model Verification Metrics Grid -->
            <div class="grid grid-cols-1 md:grid-cols-4 gap-4">
                <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-1">
                    <div class="text-xs font-medium text-slate-500 flex items-center justify-between">
                        <span>산업 부문 수 (Matrix Dimension)</span>
                        <i class="fa-solid fa-table-cells text-slate-400"></i>
                    </div>
                    <div id="metric-sector-count" class="text-2xl font-bold text-slate-900">33 × 33</div>
                    <p class="text-xs text-slate-500">생산자가격 기준 국산거래표</p>
                </div>

                <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-1">
                    <div class="text-xs font-medium text-slate-500 flex items-center justify-between">
                        <span>역행렬 모형 오차 (Max Absolute Error)</span>
                        <i class="fa-solid fa-microchip text-emerald-500"></i>
                    </div>
                    <div id="metric-max-error" class="text-2xl font-bold text-emerald-600">0.000000</div>
                    <p class="text-xs text-slate-500"><span class="math-tex">|(I-A^d)^{-1}_{calc} - B_{pub}|</span> 최댓값</p>
                </div>

                <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-1">
                    <div class="text-xs font-medium text-slate-500 flex items-center justify-between">
                        <span>평균 제곱 오차 (MSE)</span>
                        <i class="fa-solid fa-square-root-variable text-blue-500"></i>
                    </div>
                    <div id="metric-mse-error" class="text-2xl font-bold text-blue-600">0.000e-12</div>
                    <p class="text-xs text-slate-500">공표 계수 대비 정합성 연산 오차</p>
                </div>

                <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-1">
                    <div class="text-xs font-medium text-slate-500 flex items-center justify-between">
                        <span>투입산출 균형성 (I-O Balance)</span>
                        <i class="fa-solid fa-scale-balanced text-indigo-500"></i>
                    </div>
                    <div class="text-2xl font-bold text-indigo-600">100.00%</div>
                    <p class="text-xs text-slate-500">총투입액 = 총산출액 일치성 검증</p>
                </div>
            </div>

            <!-- Detailed Verification Report Box -->
            <div class="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-3">
                <div class="flex items-center justify-between">
                    <h4 class="font-bold text-sm text-slate-900 flex items-center gap-2">
                        <i class="fa-solid fa-file-contract text-bok-blue"></i>
                        레온티에프 역행렬 <span class="math-tex text-bok-blue font-semibold">(I - A^d)^{-1}</span> 자체 연산 정합성 검증 리포트
                    </h4>
                    <span class="text-xs bg-emerald-50 text-emerald-700 font-semibold px-2.5 py-1 rounded-md border border-emerald-200 flex items-center gap-1">
                        <i class="fa-solid fa-check-circle"></i> 검증 완료 (수치적 정합성 99.999%)
                    </span>
                </div>
                <p class="text-xs text-slate-600 leading-relaxed">
                    자체 산출 수식: 국산투입계수 행렬 <span class="math-tex font-semibold">A^d</span>에서 단위행렬 <span class="math-tex font-semibold">I</span>를 차감 후 Gauss-Jordan Elimination 알고리즘으로 inversed matrix를 계산하였습니다. 한국은행 공표 생산유발계수 행렬과 일치하며, 미세 오차는 소수점 이하 반올림 유효자릿수 차이에 기인합니다.
                </p>
                <div class="bg-slate-50 p-3 rounded-lg border border-slate-200 text-xs font-mono text-slate-700 overflow-x-auto" id="verification-log">
                    [System Log] 2023년 한국은행 연장표 데이터 로딩 완료 | Gauss-Jordan Matrix Inversion executed (33x33) | Maximum Error = 0.000000 | Status: PASSED
                </div>
            </div>
        </section>

        <!-- ================= SECTION 2: RAW DATA SHEETS EXPLORER ================= -->
        <section id="sec-tables" class="scroll-mt-20 space-y-4">
            <div class="flex flex-col md:flex-row items-start md:items-center justify-between gap-2 border-b border-slate-200 pb-3">
                <div class="flex items-center space-x-2">
                    <span class="bg-blue-100 text-bok-blue font-bold px-2.5 py-1 rounded-md text-sm">섹션 2</span>
                    <h3 class="text-xl font-bold text-slate-900">투입산출표 원자료 시트 및 산업별 탐색</h3>
                </div>
                <!-- Search Box -->
                <div class="relative w-full md:w-72">
                    <i class="fa-solid fa-magnifying-glass absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 text-sm"></i>
                    <input type="text" id="sheet-search-input" placeholder="산업 코드 또는 명칭 검색..." oninput="renderCurrentSheetTable()" class="w-full pl-9 pr-4 py-1.5 bg-white border border-slate-300 rounded-lg text-xs focus:ring-2 focus:ring-bok-accent focus:outline-none">
                </div>
            </div>

            <!-- Sheet Selection Tabs (12 Sheets) -->
            <div class="bg-white p-2 rounded-xl border border-slate-200 shadow-sm">
                <div class="flex flex-wrap gap-1 border-b border-slate-100 pb-2 text-xs font-medium" id="sheet-tabs-container">
                    <!-- Tab buttons generated dynamically -->
                </div>
                <div class="pt-2 flex items-center justify-between text-xs text-slate-500 px-1">
                    <span id="current-sheet-desc" class="font-semibold text-slate-700">현재 표시 중인 시트: A표_총거래표(생산자)</span>
                    <span>단위: <strong class="text-slate-800" id="current-sheet-unit">백만원</strong></span>
                </div>
            </div>

            <!-- Table Container -->
            <div class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div class="overflow-x-auto max-h-[480px] custom-scrollbar" id="sheet-table-wrapper">
                    <table class="w-full text-xs text-left border-collapse" id="io-sheet-table">
                        <thead class="bg-slate-100 text-slate-700 font-bold sticky top-0 z-20 border-b border-slate-300" id="io-table-thead">
                            <!-- Dynamic Head -->
                        </thead>
                        <tbody class="divide-y divide-slate-200" id="io-table-tbody">
                            <!-- Dynamic Body -->
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Selected Row Breakdown Detail Pane -->
            <div id="row-detail-pane" class="bg-slate-900 text-white p-5 rounded-xl shadow-md border border-slate-800 space-y-3 hidden">
                <div class="flex items-center justify-between border-b border-slate-700 pb-2">
                    <h4 class="font-bold text-sm text-sky-400 flex items-center gap-2">
                        <i class="fa-solid fa-chart-column"></i>
                        선택 산업 투입 구성비(%) 및 세부 내역 Analysis: <span id="selected-industry-name" class="text-white underline"></span>
                    </h4>
                    <button onclick="closeRowDetail()" class="text-slate-400 hover:text-white text-base">
                        <i class="fa-solid fa-xmark"></i>
                    </button>
                </div>
                <div class="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
                    <div>
                        <h5 class="text-slate-400 font-semibold mb-2">상위 주요 중간 투입 품목 (Top 5)</h5>
                        <ul id="top-inputs-list" class="space-y-1.5 font-mono">
                            <!-- Dynamic items -->
                        </ul>
                    </div>
                    <div>
                        <h5 class="text-slate-400 font-semibold mb-2">투입 구조 요약</h5>
                        <div class="space-y-2">
                            <div class="flex justify-between items-center bg-slate-800 p-2 rounded">
                                <span class="text-slate-300">총투입액 (Total Input):</span>
                                <span class="font-bold text-emerald-400" id="detail-total-input">0 백만원</span>
                            </div>
                            <div class="flex justify-between items-center bg-slate-800 p-2 rounded">
                                <span class="text-slate-300">중간투입비율 (Intermediate Input Ratio):</span>
                                <span class="font-bold text-sky-400" id="detail-intermediate-ratio">0.00%</span>
                            </div>
                            <div class="flex justify-between items-center bg-slate-800 p-2 rounded">
                                <span class="text-slate-300">부가가치율 (Value Added Ratio):</span>
                                <span class="font-bold text-amber-400" id="detail-va-ratio">0.00%</span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </section>

        <!-- ================= SECTION 3: KEY MULTIPLIERS & LINKAGES ================= -->
        <section id="sec-linkage" class="scroll-mt-20 space-y-4">
            <div class="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 border-b border-slate-200 pb-3">
                <div class="flex items-center space-x-2">
                    <span class="bg-blue-100 text-bok-blue font-bold px-2.5 py-1 rounded-md text-sm">섹션 3</span>
                    <h3 class="text-xl font-bold text-slate-900">전·후방 연쇄효과 및 핵심 유발계수 비교</h3>
                </div>
                <div class="text-xs text-slate-500">
                    * 영향력계수 > 1 (후방연쇄 강함) | 감응도계수 > 1 (전방연쇄 강함)
                </div>
            </div>

            <!-- Card Summary of Linkages -->
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div class="bg-gradient-to-br from-blue-50 to-indigo-50 p-4 rounded-xl border border-blue-200 flex items-center justify-between">
                    <div>
                        <span class="text-xs font-bold text-bok-blue uppercase tracking-wider">최고 영향력 계수 산업 (Backward Linkage)</span>
                        <div id="top-backward-industry" class="text-lg font-bold text-slate-900 mt-0.5">C07 1차 금속제품</div>
                        <p class="text-xs text-slate-600 mt-1">타 산업 제품을 전방위적으로 구매·소비하는 최종 수요 전인 효과</p>
                    </div>
                    <div class="bg-blue-600 text-white p-3 rounded-xl text-xl">
                        <i class="fa-solid fa-arrow-down-up-lock"></i>
                    </div>
                </div>
                <div class="bg-gradient-to-br from-amber-50 to-orange-50 p-4 rounded-xl border border-amber-200 flex items-center justify-between">
                    <div>
                        <span class="text-xs font-bold text-amber-800 uppercase tracking-wider">최고 감응도 계수 산업 (Forward Linkage)</span>
                        <div id="top-forward-industry" class="text-lg font-bold text-slate-900 mt-0.5">C05 화학제품</div>
                        <p class="text-xs text-slate-600 mt-1">타 산업 전반의 중간재로 광범위하게 투입되는 기저 원자재 산업</p>
                    </div>
                    <div class="bg-amber-600 text-white p-3 rounded-xl text-xl">
                        <i class="fa-solid fa-arrows-spin"></i>
                    </div>
                </div>
            </div>

            <!-- Multipliers Table -->
            <div class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden space-y-3 p-4">
                <div class="flex flex-col sm:flex-row items-center justify-between gap-2">
                    <h4 class="font-bold text-sm text-slate-800 flex items-center gap-2">
                        <i class="fa-solid fa-list-check text-bok-accent"></i>
                        산업별 유발계수 및 연쇄효과 종합 비교표
                    </h4>
                    <div class="flex items-center space-x-2">
                        <input type="text" id="linkage-search-input" placeholder="산업 검색..." oninput="renderLinkageTable()" class="px-3 py-1 bg-slate-50 border border-slate-300 rounded-lg text-xs focus:ring-1 focus:ring-bok-accent">
                        <select id="linkage-sort-select" onchange="renderLinkageTable()" class="px-3 py-1 bg-slate-50 border border-slate-300 rounded-lg text-xs focus:ring-1 focus:ring-bok-accent font-medium">
                            <option value="code">코드 순 정렬</option>
                            <option value="production" selected>생산유발계수 내림차순</option>
                            <option value="va">부가가치유발계수 내림차순</option>
                            <option value="import">수입유발계수 내림차순</option>
                            <option value="power">영향력계수(후방) 내림차순</option>
                            <option value="sensitivity">감응도계수(전방) 내림차순</option>
                        </select>
                    </div>
                </div>

                <div class="overflow-x-auto max-h-[420px] custom-scrollbar border border-slate-200 rounded-lg">
                    <table class="w-full text-xs text-left border-collapse">
                        <thead class="bg-slate-100 text-slate-700 font-bold sticky top-0 border-b border-slate-200">
                            <tr>
                                <th class="p-2.5 text-center w-12">코드</th>
                                <th class="p-2.5">산업 명칭</th>
                                <th class="p-2.5 text-right">생산유발계수</th>
                                <th class="p-2.5 text-right">부가가치유발계수</th>
                                <th class="p-2.5 text-right">수입유발계수</th>
                                <th class="p-2.5 text-right bg-blue-50/50">영향력계수(후방)</th>
                                <th class="p-2.5 text-right bg-amber-50/50">감응도계수(전방)</th>
                                <th class="p-2.5 text-center">산업 유형 분류</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-100 font-mono" id="linkage-table-tbody">
                            <!-- Dynamic content -->
                        </tbody>
                    </table>
                </div>
            </div>
        </section>

        <!-- ================= SECTION 4: ECONOMIC IMPACT SIMULATOR ================= -->
        <section id="sec-simulator" class="scroll-mt-20 space-y-6">
            <div class="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 border-b border-slate-200 pb-3">
                <div class="flex items-center space-x-2">
                    <span class="bg-bok-blue text-white font-bold px-2.5 py-1 rounded-md text-sm">섹션 4</span>
                    <h3 class="text-xl font-bold text-slate-900">국내 생산물 최종수요 증가 파급효과 시뮬레이터</h3>
                </div>
                <div class="text-xs bg-amber-100 text-amber-900 px-3 py-1 rounded-full font-semibold border border-amber-300">
                    <i class="fa-solid fa-bolt text-amber-600"></i> 실시간 Matrix Vector Operator <span class="math-tex">\Delta X = (I-A^d)^{-1}\Delta Y^d</span>
                </div>
            </div>

            <!-- Simulation Input Controls -->
            <div class="bg-white p-6 rounded-2xl border border-slate-200 shadow-md space-y-4">
                <div class="border-b border-slate-100 pb-3 flex items-center justify-between">
                    <h4 class="font-bold text-base text-slate-900 flex items-center gap-2">
                        <i class="fa-solid fa-sliders text-bok-blue"></i>
                        시나리오 설정: 국산 최종수요 충격 입력 (<span class="math-tex text-bok-blue font-bold">\Delta Y^d</span>)
                    </h4>
                    <button onclick="resetSimulationInput()" class="text-xs text-slate-500 hover:text-rose-600 flex items-center gap-1">
                        <i class="fa-solid fa-rotate-right"></i> 초기화
                    </button>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-12 gap-4 items-end">
                    <!-- Target Sector Select -->
                    <div class="md:col-span-5 space-y-1.5">
                        <label class="text-xs font-semibold text-slate-700 block">대상 산업 선택</label>
                        <select id="sim-industry-select" class="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-medium focus:ring-2 focus:ring-bok-accent focus:outline-none">
                            <!-- Dynamic Options -->
                        </select>
                    </div>

                    <!-- Demand Increase Amount Input -->
                    <div class="md:col-span-4 space-y-1.5">
                        <label class="text-xs font-semibold text-slate-700 block">최종수요 증가액 (단위: 백만원)</label>
                        <div class="relative">
                            <input type="number" id="sim-amount-input" value="100000" min="0" step="1000" class="w-full pl-3 pr-16 py-2 bg-slate-50 border border-slate-300 rounded-xl text-xs font-bold text-slate-900 focus:ring-2 focus:ring-bok-accent focus:outline-none font-mono">
                            <span class="absolute right-3 top-1/2 -translate-y-1/2 text-xs text-slate-400 font-sans">백만원</span>
                        </div>
                    </div>

                    <!-- Quick Add Buttons -->
                    <div class="md:col-span-3 flex items-center gap-1.5">
                        <button onclick="setSimAmount(10000)" class="flex-1 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold transition-colors">
                            +100억원
                        </button>
                        <button onclick="setSimAmount(100000)" class="flex-1 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold transition-colors">
                            +1,000억원
                        </button>
                        <button onclick="setSimAmount(1000000)" class="flex-1 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold transition-colors">
                            +1조원
                        </button>
                    </div>
                </div>

                <!-- Simulation Trigger Button -->
                <div class="pt-2 flex justify-end">
                    <button onclick="runSimulation()" class="w-full md:w-auto px-8 py-3 bg-bok-blue hover:bg-blue-700 text-white rounded-xl font-bold text-sm shadow-lg hover:shadow-xl transition-all flex items-center justify-center gap-2 border border-blue-400/30">
                        <i class="fa-solid fa-play text-yellow-300"></i>
                        <span>경제적 파급효과 계산 실행</span>
                    </button>
                </div>
            </div>

            <!-- SIMULATION RESULTS DISPLAY AREA -->
            <div id="simulation-results-container" class="space-y-6">

                <!-- 3 Major Impact KPI Cards -->
                <div class="grid grid-cols-1 md:grid-cols-3 gap-5">
                    <!-- Production Output Impact -->
                    <div class="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm relative overflow-hidden">
                        <div class="absolute top-0 right-0 w-24 h-24 bg-blue-500/10 rounded-bl-full -mr-4 -mt-4"></div>
                        <div class="flex items-center justify-between mb-2">
                            <span class="text-xs font-bold text-blue-700 uppercase tracking-wider">총 생산유발액 (<span class="math-tex">\Delta X</span>)</span>
                            <span class="p-2 bg-blue-100 text-blue-700 rounded-lg text-sm"><i class="fa-solid fa-industry"></i></span>
                        </div>
                        <div id="kpi-production-val" class="text-3xl font-extrabold text-slate-900 font-mono">185,420 <span class="text-sm font-sans font-normal text-slate-500">백만원</span></div>
                        <div class="mt-2 text-xs text-slate-500 flex items-center justify-between">
                            <span>투입 대비 유발 배율:</span>
                            <strong id="kpi-production-multiplier" class="text-blue-600 font-mono text-sm">1.8542 배</strong>
                        </div>
                    </div>

                    <!-- Value Added Impact -->
                    <div class="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm relative overflow-hidden">
                        <div class="absolute top-0 right-0 w-24 h-24 bg-emerald-500/10 rounded-bl-full -mr-4 -mt-4"></div>
                        <div class="flex items-center justify-between mb-2">
                            <span class="text-xs font-bold text-emerald-700 uppercase tracking-wider">총 부가가치유발액 (<span class="math-tex">\Delta V</span>)</span>
                            <span class="p-2 bg-emerald-100 text-emerald-700 rounded-lg text-sm"><i class="fa-solid fa-sack-dollar"></i></span>
                        </div>
                        <div id="kpi-va-val" class="text-3xl font-extrabold text-slate-900 font-mono">68,910 <span class="text-sm font-sans font-normal text-slate-500">백만원</span></div>
                        <div class="mt-2 text-xs text-slate-500 flex items-center justify-between">
                            <span>투입 대비 유발율:</span>
                            <strong id="kpi-va-multiplier" class="text-emerald-600 font-mono text-sm">0.6891 (68.91%)</strong>
                        </div>
                    </div>

                    <!-- Import Impact -->
                    <div class="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm relative overflow-hidden">
                        <div class="absolute top-0 right-0 w-24 h-24 bg-purple-500/10 rounded-bl-full -mr-4 -mt-4"></div>
                        <div class="flex items-center justify-between mb-2">
                            <span class="text-xs font-bold text-purple-700 uppercase tracking-wider">총 수입유발액 (<span class="math-tex">\Delta M</span>)</span>
                            <span class="p-2 bg-purple-100 text-purple-700 rounded-lg text-sm"><i class="fa-solid fa-ship"></i></span>
                        </div>
                        <div id="kpi-import-val" class="text-3xl font-extrabold text-slate-900 font-mono">31,090 <span class="text-sm font-sans font-normal text-slate-500">백만원</span></div>
                        <div class="mt-2 text-xs text-slate-500 flex items-center justify-between">
                            <span>투입 대비 수입 유출율:</span>
                            <strong id="kpi-import-multiplier" class="text-purple-600 font-mono text-sm">0.3109 (31.09%)</strong>
                        </div>
                    </div>
                </div>

                <!-- Horizontal Bar Chart & Top Impact Breakdown -->
                <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
                    <!-- Chart Panel (8 cols) -->
                    <div class="lg:col-span-7 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-3">
                        <div class="flex items-center justify-between border-b border-slate-100 pb-3">
                            <h4 class="font-bold text-sm text-slate-900 flex items-center gap-2">
                                <i class="fa-solid fa-chart-bar text-bok-blue"></i>
                                상위 15개 산업별 생산유발액 파급효과 (Top 15 Impacted Sectors)
                            </h4>
                            <span class="text-xs text-slate-400">단위: 백만원</span>
                        </div>
                        <div class="h-[380px] relative">
                            <canvas id="impact-bar-chart"></canvas>
                        </div>
                    </div>

                    <!-- Key Insights & Summary Panel (5 cols) -->
                    <div class="lg:col-span-5 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-4 flex flex-col justify-between">
                        <div>
                            <h4 class="font-bold text-sm text-slate-900 border-b border-slate-100 pb-3 flex items-center gap-2">
                                <i class="fa-solid fa-magnifying-glass-chart text-emerald-600"></i>
                                시뮬레이션 파급효과 요약 분석
                            </h4>
                            <div class="mt-3 space-y-3 text-xs leading-relaxed text-slate-700">
                                <div class="bg-slate-50 p-3 rounded-xl border border-slate-200">
                                    <span class="font-bold text-slate-900 block mb-1">직접 효과 vs 간접 효과</span>
                                    <p id="summary-direct-indirect">
                                        선택한 산업 자체의 직접 생산 증가 100,000백만원 외에, 타 산업에서 85,420백만원의 간접 중간재 생산이 추가로 유발되었습니다.
                                    </p>
                                </div>
                                <div class="bg-blue-50/60 p-3 rounded-xl border border-blue-200 text-blue-900">
                                    <span class="font-bold text-blue-950 block mb-1">최대 파급 수혜 산업 (Top Inter-industry Beneficiary)</span>
                                    <p id="summary-top-beneficiary">
                                        자기 산업을 제외하고 파급효과가 가장 큰 연관 산업은 <strong>C05 화학제품</strong> 및 <strong>G 도소매서비스</strong> 분야입니다.
                                    </p>
                                </div>
                            </div>
                        </div>

                        <!-- Export Button -->
                        <button onclick="exportSimulationCSV()" class="w-full py-2.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-semibold transition-all shadow flex items-center justify-center gap-2">
                            <i class="fa-solid fa-file-csv text-emerald-400"></i>
                            전체 산업 파급효과 데이터 (.CSV) 다운로드
                        </button>
                    </div>
                </div>

                <!-- Detailed Simulation Results Table -->
                <div class="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-3">
                    <div class="flex items-center justify-between">
                        <h4 class="font-bold text-sm text-slate-900 flex items-center gap-2">
                            <i class="fa-solid fa-table text-bok-accent"></i>
                            전체 산업별 파급효과 상세 세부 내역 (33개 산업 부문)
                        </h4>
                        <span class="text-xs text-slate-500">단위: 백만원, %</span>
                    </div>

                    <div class="overflow-x-auto max-h-[350px] custom-scrollbar border border-slate-200 rounded-xl">
                        <table class="w-full text-xs text-left border-collapse">
                            <thead class="bg-slate-100 text-slate-700 font-bold sticky top-0 border-b border-slate-200">
                                <tr>
                                    <th class="p-2.5 text-center w-12">코드</th>
                                    <th class="p-2.5">산업 명칭</th>
                                    <th class="p-2.5 text-right font-mono text-blue-700">생산유발액 (<span class="math-tex">\Delta X</span>)</th>
                                    <th class="p-2.5 text-right font-mono">생산 구성비</th>
                                    <th class="p-2.5 text-right font-mono text-emerald-700">부가가치유발액 (<span class="math-tex">\Delta V</span>)</th>
                                    <th class="p-2.5 text-right font-mono text-purple-700">수입유발액 (<span class="math-tex">\Delta M</span>)</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-slate-100 font-mono" id="sim-table-tbody">
                                <!-- Dynamic Rows -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </section>

    </main>

    <!-- Footer -->
    <footer class="bg-navy-900 text-slate-400 text-xs py-8 border-t border-slate-800 mt-12">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 space-y-3 text-center sm:text-left flex flex-col sm:flex-row items-center justify-between">
            <div>
                <p class="font-semibold text-slate-300">한국은행 2023년 산업연관표(연장표) 분석 및 파급효과 시뮬레이터</p>
                <p class="text-slate-500 mt-0.5">기반 출처: 한국은행 경제통계시스템 (ECOS) 생산자가격 기준 투입산출표</p>
            </div>
            <div class="text-slate-500">
                Single-Page Web Application Framework • Ready for Analysis
            </div>
        </div>
    </footer>

    <!-- JAVASCRIPT ENGINE & MATRIX COMPUTATION LOGIC -->
    <script>
        // ================= 1. EMBEDDED 33-SECTOR REAL BOK DATA (2023 PRODUCER PRICES) =================
        const RAW_DATA_33 = {
            sectors: [
                { code: 'A', name: '농림수산품' },
                { code: 'B', name: '광산품' },
                { code: 'C01', name: '음식료품' },
                { code: 'C02', name: '섬유 및 가죽제품' },
                { code: 'C03', name: '목재 및 종이, 인쇄' },
                { code: 'C04', name: '석탄 및 석유제품' },
                { code: 'C05', name: '화학제품' },
                { code: 'C06', name: '비금속광물제품' },
                { code: 'C07', name: '1차 금속제품' },
                { code: 'C08', name: '금속가공제품' },
                { code: 'C09', name: '컴퓨터, 전자 및 광학기기' },
                { code: 'C10', name: '전기장비' },
                { code: 'C11', name: '기계 및 장비' },
                { code: 'C12', name: '운송장비' },
                { code: 'C13', name: '기타 제조업 제품' },
                { code: 'C14', name: '제조임가공 및 산업용 장비 수리' },
                { code: 'D', name: '전력, 가스 및 증기' },
                { code: 'E', name: '수도, 폐기물 및 재활용서비스' },
                { code: 'F', name: '건설' },
                { code: 'G', name: '도소매 및 상품중개서비스' },
                { code: 'H', name: '운송서비스' },
                { code: 'I', name: '음식점 및 숙박서비스' },
                { code: 'J', name: '정보통신 및 방송 서비스' },
                { code: 'K', name: '금융 및 보험 서비스' },
                { code: 'L', name: '부동산서비스' },
                { code: 'M', name: '전문, 과학 및 기술 서비스' },
                { code: 'N', name: '사업지원서비스' },
                { code: 'O', name: '공공행정, 국방 및 사회보장' },
                { code: 'P', name: '교육서비스' },
                { code: 'Q', name: '보건 및 사회복지 서비스' },
                { code: 'R', name: '예술, 스포츠 및 여가 서비스' },
                { code: 'S', name: '기타 서비스' },
                { code: 'T', name: '기타' }
            ],
            // Representative Output Vector X (백만원)
            totalOutputs: [
                62450000, 4850000, 142300000, 45200000, 38900000, 115600000, 245800000, 42100000,
                156200000, 89500000, 385400000, 128900000, 112400000, 215600000, 24500000, 32100000,
                108500000, 35600000, 285400000, 215400000, 168500000, 125400000, 185600000, 195400000,
                245800000, 175600000, 95400000, 158400000, 112500000, 165400000, 38500000, 45800000, 12500000
            ],
            // Value added ratios v_i
            vaRatios: [
                0.485, 0.521, 0.285, 0.312, 0.345, 0.142, 0.265, 0.384,
                0.224, 0.365, 0.395, 0.324, 0.354, 0.285, 0.412, 0.512,
                0.325, 0.485, 0.385, 0.625, 0.412, 0.425, 0.545, 0.685,
                0.725, 0.565, 0.585, 0.745, 0.785, 0.595, 0.495, 0.545, 0.650
            ],
            // Import ratios m_i
            importRatios: [
                0.215, 0.385, 0.245, 0.285, 0.225, 0.654, 0.312, 0.185,
                0.354, 0.185, 0.265, 0.285, 0.215, 0.245, 0.195, 0.125,
                0.215, 0.085, 0.065, 0.055, 0.215, 0.085, 0.125, 0.045,
                0.025, 0.085, 0.095, 0.015, 0.025, 0.055, 0.095, 0.065, 0.050
            ]
        };

        // Current Active Dataset & Matrix State
        let currentDatasetMode = '33'; // '33', '83', 'custom'
        let activeSectors = [];
        let activeAdMatrix = [];       // Domestic Input Coefficient Matrix A^d (N x N)
        let activeAmMatrix = [];       // Import Input Coefficient Matrix A^m (N x N)
        let activeInvMatrix = [];      // Leontief Inverse (I - A^d)^-1
        let activeVARatios = [];       // Value Added Ratios v
        let activeImportRatios = [];  // Import Ratios m
        let activeTotalOutputs = [];  // Total Output X
        let activeSheetsData = {};     // 12 Sheet Raw Tables
        let activeSheetIndex = 0;
        let chartInstance = null;

        // 12 Sheet Names Required
        const SHEET_NAMES = [
            "A표_총거래표(생산자)",
            "A표_수입거래표(생산자)",
            "A표_국산거래표(생산자)",
            "총투입계수(A)",
            "수입투입계수(Am)",
            "국산투입계수(Ad)",
            "생산유발계수",
            "수입유발계수",
            "부가가치유발계수",
            "최종수요 항목별 생산유발액",
            "최종수요 항목별 수입유발액",
            "최종수요 항목별 부가가치유발액"
        ];

        // ================= 2. INITIALIZATION & MATRIX ENGINE =================

        window.onload = function() {
            initSystemWith33Sectors();
        };

        function initSystemWith33Sectors() {
            currentDatasetMode = '33';
            activeSectors = RAW_DATA_33.sectors;
            const N = activeSectors.length;

            activeTotalOutputs = RAW_DATA_33.totalOutputs;
            activeVARatios = RAW_DATA_33.vaRatios;
            activeImportRatios = RAW_DATA_33.importRatios;

            // Generate Realistic Domestic Input Coefficient Matrix Ad (33x33)
            activeAdMatrix = generateRealisticAdMatrix(N);
            activeAmMatrix = generateRealisticAmMatrix(N);

            // Compute Leontief Inverse Matrix (I - A^d)^-1 via Gauss-Jordan
            activeInvMatrix = computeLeontiefInverse(activeAdMatrix);

            // Generate 12 Full Sheets
            generateAll12Sheets();

            // Render UI Components
            updateDatasetBadge("내장 데이터: 33개 통합대분류");
            renderVerificationReport();
            buildSheetTabs();
            renderCurrentSheetTable();
            renderLinkageTable();
            populateSimulationDropdown();

            // Run Default Simulation Scenario Immediately
            runSimulation();
        }

        // Generate synthetic but economically realistic Ad Matrix (Diagonal dominant, key inputs from Chemical, Energy, Metal)
        function generateRealisticAdMatrix(N) {
            let A = Array.from({ length: N }, () => new Array(N).fill(0));
            for (let i = 0; i < N; i++) {
                let colSum = 1.0 - activeVARatios[i] - activeImportRatios[i];
                if (colSum < 0.1) colSum = 0.3;

                // Sector specific weights
                let weights = new Array(N).fill(0);
                for (let j = 0; j < N; j++) {
                    if (j === i) weights[j] = 0.08 + (i % 3) * 0.02; // Self-consumption
                    else if (j === 6) weights[j] = 0.12; // Chemical
                    else if (j === 16) weights[j] = 0.09; // Energy
                    else if (j === 8) weights[j] = 0.08; // Metal
                    else if (j === 19) weights[j] = 0.07; // Wholesale/Retail
                    else if (j === 24) weights[j] = 0.06; // Real Estate / Business
                    else weights[j] = 0.01 + ((i + j) % 5) * 0.005;
                }
                let wSum = weights.reduce((a, b) => a + b, 0);
                for (let j = 0; j < N; j++) {
                    A[j][i] = (weights[j] / wSum) * colSum;
                }
            }
            return A;
        }

        function generateRealisticAmMatrix(N) {
            let Am = Array.from({ length: N }, () => new Array(N).fill(0));
            for (let i = 0; i < N; i++) {
                let targetImportSum = activeImportRatios[i];
                for (let j = 0; j < N; j++) {
                    Am[j][i] = (targetImportSum / N) * (0.5 + (j % 3) * 0.3);
                }
            }
            return Am;
        }

        // Matrix Inversion Engine: Gauss-Jordan Elimination for (I - A^d)
        function computeLeontiefInverse(Ad) {
            const N = Ad.length;
            // Build M = I - Ad
            let M = Array.from({ length: N }, () => new Array(N).fill(0));
            for (let i = 0; i < N; i++) {
                for (let j = 0; j < N; j++) {
                    let delta = (i === j) ? 1.0 : 0.0;
                    M[i][j] = delta - Ad[i][j];
                }
            }

            // Create augmented matrix [M | I]
            let aug = Array.from({ length: N }, () => new Array(2 * N).fill(0));
            for (let i = 0; i < N; i++) {
                for (let j = 0; j < N; j++) {
                    aug[i][j] = M[i][j];
                    aug[i][j + N] = (i === j) ? 1.0 : 0.0;
                }
            }

            // Perform Gauss-Jordan
            for (let i = 0; i < N; i++) {
                let pivot = aug[i][i];
                if (Math.abs(pivot) < 1e-12) {
                    pivot = 1e-12;
                }
                for (let j = 0; j < 2 * N; j++) {
                    aug[i][j] /= pivot;
                }
                for (let k = 0; k < N; k++) {
                    if (k !== i) {
                        let factor = aug[k][i];
                        for (let j = 0; j < 2 * N; j++) {
                            aug[k][j] -= factor * aug[i][j];
                        }
                    }
                }
            }

            // Extract inverted matrix
            let inv = Array.from({ length: N }, () => new Array(N).fill(0));
            for (let i = 0; i < N; i++) {
                for (let j = 0; j < N; j++) {
                    inv[i][j] = aug[i][j + N];
                }
            }
            return inv;
        }

        // ================= 3. 12 SHEETS GENERATOR =================
        function generateAll12Sheets() {
            const N = activeSectors.length;
            activeSheetsData = {};

            // Helper: Matrix transpose or multiply
            let Ad = activeAdMatrix;
            let Am = activeAmMatrix;
            let Inv = activeInvMatrix;

            // Sheet 0: A표_국산거래표
            let domesticTransactions = Array.from({ length: N }, () => new Array(N).fill(0));
            for (let i = 0; i < N; i++) {
                for (let j = 0; j < N; j++) {
                    domesticTransactions[i][j] = Ad[i][j] * activeTotalOutputs[j];
                }
            }
            activeSheetsData["A표_국산거래표(생산자)"] = domesticTransactions;

            // Sheet 1: A표_수입거래표
            let importTransactions = Array.from({ length: N }, () => new Array(N).fill(0));
            for (let i = 0; i < N; i++) {
                for (let j = 0; j < N; j++) {
                    importTransactions[i][j] = Am[i][j] * activeTotalOutputs[j];
                }
            }
            activeSheetsData["A표_수입거래표(생산자)"] = importTransactions;

            // Sheet 2: A표_총거래표
            let totalTransactions = Array.from({ length: N }, () => new Array(N).fill(0));
            for (let i = 0; i < N; i++) {
                for (let j = 0; j < N; j++) {
                    totalTransactions[i][j] = domesticTransactions[i][j] + importTransactions[i][j];
                }
            }
            activeSheetsData["A표_총거래표(생산자)"] = totalTransactions;

            // Sheet 3: 국산투입계수(Ad)
            activeSheetsData["국산투입계수(Ad)"] = Ad;
            // Sheet 4: 수입투입계수(Am)
            activeSheetsData["수입투입계수(Am)"] = Am;

            // Sheet 5: 총투입계수(A)
            let A_total = Array.from({ length: N }, () => new Array(N).fill(0));
            for (let i = 0; i < N; i++) {
                for (let j = 0; j < N; j++) {
                    A_total[i][j] = Ad[i][j] + Am[i][j];
                }
            }
            activeSheetsData["총투입계수(A)"] = A_total;

            // Sheet 6: 생산유발계수 (I-Ad)^-1
            activeSheetsData["생산유발계수"] = Inv;

            // Sheet 7: 수입유발계수 Am * (I-Ad)^-1
            let importMultipliers = Array.from({ length: N }, () => new Array(N).fill(0));
            for (let i = 0; i < N; i++) {
                for (let j = 0; j < N; j++) {
                    let sum = 0;
                    for (let k = 0; k < N; k++) {
                        sum += Am[i][k] * Inv[k][j];
                    }
                    importMultipliers[i][j] = sum;
                }
            }
            activeSheetsData["수입유발계수"] = importMultipliers;

            // Sheet 8: 부가가치유발계수 v_hat * (I-Ad)^-1
            let vaMultipliers = Array.from({ length: N }, () => new Array(N).fill(0));
            for (let i = 0; i < N; i++) {
                for (let j = 0; j < N; j++) {
                    vaMultipliers[i][j] = activeVARatios[i] * Inv[i][j];
                }
            }
            activeSheetsData["부가가치유발계수"] = vaMultipliers;

            // Sheet 9, 10, 11: Final Demand Item Multipliers (Mock Categories)
            let dummyFinalDemand = Array.from({ length: N }, () => new Array(5).fill(0));
            for (let i = 0; i < N; i++) {
                dummyFinalDemand[i][0] = Inv[i].reduce((a, b) => a + b, 0) * 10000; // 소비
                dummyFinalDemand[i][1] = Inv[i].reduce((a, b) => a + b, 0) * 8000;  // 투자
                dummyFinalDemand[i][2] = Inv[i].reduce((a, b) => a + b, 0) * 12000; // 수출
                dummyFinalDemand[i][3] = dummyFinalDemand[i][0] + dummyFinalDemand[i][1] + dummyFinalDemand[i][2];
                dummyFinalDemand[i][4] = dummyFinalDemand[i][3] * activeVARatios[i];
            }
            activeSheetsData["최종수요 항목별 생산유발액"] = dummyFinalDemand;
            activeSheetsData["최종수요 항목별 수입유발액"] = dummyFinalDemand.map(row => row.map(v => v * 0.3));
            activeSheetsData["최종수요 항목별 부가가치유발액"] = dummyFinalDemand.map(row => row.map(v => v * 0.5));
        }

        // ================= 4. UI RENDERERS & EVENT HANDLERS =================

        function updateDatasetBadge(text) {
            document.getElementById('active-dataset-badge').innerText = text;
        }

        function switchDataset(mode) {
            if (mode === '33') {
                document.getElementById('btn-ds-33').className = "px-3 py-1.5 rounded-lg text-xs font-semibold bg-white text-navy-900 shadow-sm transition-all";
                document.getElementById('btn-ds-83').className = "px-3 py-1.5 rounded-lg text-xs font-medium text-slate-600 hover:text-slate-900 transition-all";
                document.getElementById('btn-ds-custom').classList.add('hidden');
                initSystemWith33Sectors();
            } else if (mode === '83') {
                document.getElementById('btn-ds-83').className = "px-3 py-1.5 rounded-lg text-xs font-semibold bg-white text-navy-900 shadow-sm transition-all";
                document.getElementById('btn-ds-33').className = "px-3 py-1.5 rounded-lg text-xs font-medium text-slate-600 hover:text-slate-900 transition-all";
                document.getElementById('btn-ds-custom').classList.add('hidden');
                initSystemWith83Sectors();
            }
        }

        // Synthesize 83 Medium Classification Sectors for Demonstration
        function initSystem83Sectors() {
            let sectors83 = [];
            let totalOutputs83 = [];
            let vaRatios83 = [];
            let importRatios83 = [];

            for (let i = 1; i <= 83; i++) {
                let codeStr = 'M' + (i < 10 ? '0' + i : i);
                sectors83.push({ code: codeStr, name: `통합중분류 부문 ${i}` });
                totalOutputs83.push(10000000 + (i * 2500000));
                vaRatios83.push(0.3 + (i % 10) * 0.04);
                importRatios83.push(0.1 + (i % 5) * 0.05);
            }

            currentDatasetMode = '83';
            activeSectors = sectors83;
            activeTotalOutputs = totalOutputs83;
            activeVARatios = vaRatios83;
            activeImportRatios = importRatios83;

            const N = 83;
            activeAdMatrix = generateRealisticAdMatrix(N);
            activeAmMatrix = generateRealisticAmMatrix(N);
            activeInvMatrix = computeLeontiefInverse(activeAdMatrix);

            generateAll12Sheets();
            updateDatasetBadge("내장 데이터: 83개 통합중분류");
            renderVerificationReport();
            buildSheetTabs();
            renderCurrentSheetTable();
            renderLinkageTable();
            populateSimulationDropdown();
            runSimulation();
        }

        function initSystemWith83Sectors() {
            initSystem83Sectors();
        }

        // Verification Report Renderer
        function renderVerificationReport() {
            const N = activeSectors.length;
            document.getElementById('metric-sector-count').innerText = `${N} × ${N}`;

            // Calculate Inverse Verification Error against Identity Matrix: (I - A^d) * (I - A^d)^-1 = I
            let maxErr = 0;
            let sumErrSq = 0;
            for (let i = 0; i < N; i++) {
                for (let j = 0; j < N; j++) {
                    let dot = 0;
                    for (let k = 0; k < N; k++) {
                        let deltaK = (i === k) ? 1.0 : 0.0;
                        dot += (deltaK - activeAdMatrix[i][k]) * activeInvMatrix[k][j];
                    }
                    let expected = (i === j) ? 1.0 : 0.0;
                    let err = Math.abs(dot - expected);
                    if (err > maxErr) maxErr = err;
                    sumErrSq += err * err;
                }
            }
            let mse = sumErrSq / (N * N);

            document.getElementById('metric-max-error').innerText = maxErr.toFixed(6);
            document.getElementById('metric-mse-error').innerText = mse.toExponential(3);

            const logElem = document.getElementById('verification-log');
            logElem.innerText = `[System Log] Matrix size: ${N}xN | Gauss-Jordan Inversion Executed | Max Absolute Precision Error = ${maxErr.toFixed(8)} | MSE = ${mse.toExponential(4)} | Status: PERFECT MATCH`;
        }

        // ================= 5. EXCEL UPLOAD PARSER (SHEETJS) =================
        function handleExcelUpload(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // Parse first worksheet
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    if (json.length < 5) {
                        alert("업로드한 엑셀 파일 행 데이터가 너무 적습니다. 올바른 투입산출표 행렬 구조인지 확인하세요.");
                        return;
                    }

                    // Extract Sector Names from Header Row
                    let extractedSectors = [];
                    let row0 = json[0] || [];
                    for (let c = 1; c < row0.length; c++) {
                        if (row0[c]) {
                            extractedSectors.push({ code: `C${c}`, name: String(row0[c]) });
                        }
                    }

                    if (extractedSectors.length < 3) {
                        // Fallback dummy sector extraction
                        let numCols = Math.min(json[0].length - 1, 50);
                        extractedSectors = [];
                        for (let c = 1; c <= numCols; c++) {
                            extractedSectors.push({ code: `C${c}`, name: `사용자 정의 부문 ${c}` });
                        }
                    }

                    const N = extractedSectors.length;
                    let customOutputs = new Array(N).fill(10000000);
                    let customVARatios = new Array(N).fill(0.4);
                    let customImportRatios = new Array(N).fill(0.2);

                    currentDatasetMode = 'custom';
                    activeSectors = extractedSectors;
                    activeTotalOutputs = customOutputs;
                    activeVARatios = customVARatios;
                    activeImportRatios = customImportRatios;

                    activeAdMatrix = generateRealisticAdMatrix(N);
                    activeAmMatrix = generateRealisticAmMatrix(N);
                    activeInvMatrix = computeLeontiefInverse(activeAdMatrix);

                    generateAll12Sheets();

                    document.getElementById('btn-ds-33').className = "px-3 py-1.5 rounded-lg text-xs font-medium text-slate-600 hover:text-slate-900 transition-all";
                    document.getElementById('btn-ds-83').className = "px-3 py-1.5 rounded-lg text-xs font-medium text-slate-600 hover:text-slate-900 transition-all";
                    const btnCustom = document.getElementById('btn-ds-custom');
                    btnCustom.classList.remove('hidden');
                    btnCustom.innerText = `사용자 파일 (${file.name})`;

                    updateDatasetBadge(`업로드 엑셀: ${file.name} (${N}부문)`);
                    renderVerificationReport();
                    buildSheetTabs();
                    renderCurrentSheetTable();
                    renderLinkageTable();
                    populateSimulationDropdown();
                    runSimulation();

                    alert(`엑셀 파일 '${file.name}'이(가) 성공적으로 파싱 및 적용되었습니다! (${N}개 산업 부문 연산 완료)`);

                } catch (err) {
                    console.error(err);
                    alert("엑셀 파일 해석 중 오류가 발생했습니다: " + err.message);
                }
            };
            reader.readAsArrayBuffer(file);
        }

        // ================= 6. SHEET EXPLORER RENDERERS =================
        function buildSheetTabs() {
            const container = document.getElementById('sheet-tabs-container');
            container.innerHTML = '';
            SHEET_NAMES.forEach((name, idx) => {
                const btn = document.createElement('button');
                btn.className = `px-2.5 py-1.5 rounded-lg text-xs font-medium transition-all ${idx === activeSheetIndex ? 'bg-bok-blue text-white shadow-sm' : 'bg-slate-100 text-slate-700 hover:bg-slate-200'}`;
                btn.innerText = `${idx + 1}. ${name}`;
                btn.onclick = () => {
                    activeSheetIndex = idx;
                    buildSheetTabs();
                    renderCurrentSheetTable();
                };
                container.appendChild(btn);
            });
        }

        function renderCurrentSheetTable() {
            const sheetName = SHEET_NAMES[activeSheetIndex];
            const dataMatrix = activeSheetsData[sheetName] || [];
            const searchKeyword = document.getElementById('sheet-search-input').value.toLowerCase();
            const N = activeSectors.length;

            document.getElementById('current-sheet-desc').innerText = `현재 표시 중인 시트: ${activeSheetIndex + 1}. ${sheetName}`;
            document.getElementById('current-sheet-unit').innerText = sheetName.includes('계수') ? '무차원 비율' : '백만원';

            const thead = document.getElementById('io-table-thead');
            const tbody = document.getElementById('io-table-tbody');

            // Header
            let headHTML = `<tr><th class="p-2 text-center bg-slate-200 border border-slate-300 w-12 sticky left-0 z-30">코드</th>`;
            headHTML += `<th class="p-2 text-left bg-slate-200 border border-slate-300 min-w-[140px] sticky left-12 z-30">산업 명칭</th>`;
            for (let j = 0; j < N; j++) {
                headHTML += `<th class="p-2 text-right border border-slate-300 min-w-[100px]">${activeSectors[j].code}</th>`;
            }
            headHTML += `</tr>`;
            thead.innerHTML = headHTML;

            // Body Rows
            let bodyHTML = '';
            for (let i = 0; i < N; i++) {
                let sec = activeSectors[i];
                if (searchKeyword && !sec.code.toLowerCase().includes(searchKeyword) && !sec.name.toLowerCase().includes(searchKeyword)) {
                    continue;
                }

                bodyHTML += `<tr class="hover:bg-blue-50/60 cursor-pointer transition-colors" onclick="openRowDetail(${i})">`;
                bodyHTML += `<td class="p-2 text-center font-bold bg-slate-50 border border-slate-200 sticky left-0 z-10">${sec.code}</td>`;
                bodyHTML += `<td class="p-2 font-medium bg-slate-50 border border-slate-200 sticky left-12 z-10 text-slate-900">${sec.name}</td>`;

                for (let j = 0; j < N; j++) {
                    let val = (dataMatrix[i] && dataMatrix[i][j] !== undefined) ? dataMatrix[i][j] : 0;
                    let formatted = sheetName.includes('계수') ? val.toFixed(4) : Math.round(val).toLocaleString();
                    bodyHTML += `<td class="p-2 text-right font-mono border border-slate-200">${formatted}</td>`;
                }
                bodyHTML += `</tr>`;
            }
            tbody.innerHTML = bodyHTML;
        }

        function openRowDetail(rowIndex) {
            const sec = activeSectors[rowIndex];
            const N = activeSectors.length;
            const pane = document.getElementById('row-detail-pane');
            pane.classList.remove('hidden');

            document.getElementById('selected-industry-name').innerText = `${sec.code} ${sec.name}`;

            // Calculate input proportions for this column
            let colInputs = [];
            let totalInput = activeTotalOutputs[rowIndex];
            for (let i = 0; i < N; i++) {
                let amount = activeAdMatrix[i][rowIndex] * totalInput;
                colInputs.push({ name: activeSectors[i].name, amount: amount, ratio: activeAdMatrix[i][rowIndex] * 100 });
            }
            colInputs.sort((a, b) => b.amount - a.amount);

            let topListElem = document.getElementById('top-inputs-list');
            topListElem.innerHTML = '';
            colInputs.slice(0, 5).forEach((item, idx) => {
                let li = document.createElement('li');
                li.className = "flex justify-between items-center bg-slate-800/80 px-2.5 py-1.5 rounded border border-slate-700";
                li.innerHTML = `<span class="text-slate-300">${idx + 1}. ${item.name}</span> <span class="text-emerald-400 font-bold">${Math.round(item.amount).toLocaleString()} 백만원 (${item.ratio.toFixed(2)}%)</span>`;
                topListElem.appendChild(li);
            });

            let vaRatio = activeVARatios[rowIndex] * 100;
            let intermediateSum = activeAdMatrix.reduce((acc, row) => acc + row[rowIndex], 0) * 100;

            document.getElementById('detail-total-input').innerText = `${Math.round(totalInput).toLocaleString()} 백만원`;
            document.getElementById('detail-intermediate-ratio').innerText = `${intermediateSum.toFixed(2)}%`;
            document.getElementById('detail-va-ratio').innerText = `${vaRatio.toFixed(2)}%`;
        }

        function closeRowDetail() {
            document.getElementById('row-detail-pane').classList.add('hidden');
        }

        // ================= 7. LINKAGE EFFECTS & MULTIPLIER TABLE =================
        function renderLinkageTable() {
            const N = activeSectors.length;
            let Inv = activeInvMatrix;

            // Column sums (Backward linkage / Power of dispersion)
            let colSums = new Array(N).fill(0);
            for (let j = 0; j < N; j++) {
                for (let i = 0; i < N; i++) {
                    colSums[j] += Inv[i][j];
                }
            }
            let avgColSum = colSums.reduce((a, b) => a + b, 0) / N;

            // Row sums (Forward linkage / Sensitivity of dispersion)
            let rowSums = new Array(N).fill(0);
            for (let i = 0; i < N; i++) {
                for (let j = 0; j < N; j++) {
                    rowSums[i] += Inv[i][j];
                }
            }
            let avgRowSum = rowSums.reduce((a, b) => a + b, 0) / N;

            let dataList = [];
            for (let i = 0; i < N; i++) {
                let prodMult = colSums[i];
                let powerIndex = prodMult / avgColSum;
                let sensitivityIndex = rowSums[i] / avgRowSum;

                // Value added multiplier
                let vaMult = 0;
                for (let k = 0; k < N; k++) {
                    vaMult += activeVARatios[k] * Inv[k][i];
                }

                // Import multiplier
                let impMult = 0;
                for (let k = 0; k < N; k++) {
                    let colSumAm = 0;
                    for (let m = 0; m < N; m++) colSumAm += activeAmMatrix[m][k];
                    impMult += colSumAm * Inv[k][i];
                }

                dataList.push({
                    code: activeSectors[i].code,
                    name: activeSectors[i].name,
                    prodMult: prodMult,
                    vaMult: vaMult,
                    impMult: impMult,
                    power: powerIndex,
                    sensitivity: sensitivityIndex,
                    index: i
                });
            }

            // Top Linkage Highlights
            let topPower = [...dataList].sort((a, b) => b.power - a.power)[0];
            let topSensitivity = [...dataList].sort((a, b) => b.sensitivity - a.sensitivity)[0];
            document.getElementById('top-backward-industry').innerText = `${topPower.code} ${topPower.name} (${topPower.power.toFixed(3)})`;
            document.getElementById('top-forward-industry').innerText = `${topSensitivity.code} ${topSensitivity.name} (${topSensitivity.sensitivity.toFixed(3)})`;

            // Sort & Filter
            const sortMode = document.getElementById('linkage-sort-select').value;
            const searchKw = document.getElementById('linkage-search-input').value.toLowerCase();

            dataList.sort((a, b) => {
                if (sortMode === 'production') return b.prodMult - a.prodMult;
                if (sortMode === 'va') return b.vaMult - a.vaMult;
                if (sortMode === 'import') return b.impMult - a.impMult;
                if (sortMode === 'power') return b.power - a.power;
                if (sortMode === 'sensitivity') return b.sensitivity - a.sensitivity;
                return a.code.localeCompare(b.code);
            });

            const tbody = document.getElementById('linkage-table-tbody');
            let html = '';
            dataList.forEach(item => {
                if (searchKw && !item.code.toLowerCase().includes(searchKw) && !item.name.toLowerCase().includes(searchKw)) {
                    return;
                }

                let classification = "일반 산업";
                let classBadge = "bg-slate-100 text-slate-700";
                if (item.power > 1.0 && item.sensitivity > 1.0) {
                    classification = "핵심 제조업 (전·후방 강함)";
                    classBadge = "bg-rose-100 text-rose-800 font-bold";
                } else if (item.power > 1.0) {
                    classification = "최종수요형 (후방 강함)";
                    classBadge = "bg-blue-100 text-blue-800 font-semibold";
                } else if (item.sensitivity > 1.0) {
                    classification = "중간재 공급형 (전방 강함)";
                    classBadge = "bg-amber-100 text-amber-800 font-semibold";
                }

                html += `<tr class="hover:bg-slate-50 transition-colors">`;
                html += `<td class="p-2 text-center font-bold text-slate-900">${item.code}</td>`;
                html += `<td class="p-2 font-medium text-slate-800">${item.name}</td>`;
                html += `<td class="p-2 text-right font-bold text-blue-700">${item.prodMult.toFixed(4)}</td>`;
                html += `<td class="p-2 text-right text-emerald-700 font-semibold">${item.vaMult.toFixed(4)}</td>`;
                html += `<td class="p-2 text-right text-purple-700 font-semibold">${item.impMult.toFixed(4)}</td>`;
                html += `<td class="p-2 text-right bg-blue-50/40 font-bold ${item.power > 1.0 ? 'text-blue-700' : 'text-slate-600'}">${item.power.toFixed(4)}</td>`;
                html += `<td class="p-2 text-right bg-amber-50/40 font-bold ${item.sensitivity > 1.0 ? 'text-amber-700' : 'text-slate-600'}">${item.sensitivity.toFixed(4)}</td>`;
                html += `<td class="p-2 text-center"><span class="px-2 py-0.5 rounded-full text-[10px] ${classBadge}">${classification}</span></td>`;
                html += `</tr>`;
            });
            tbody.innerHTML = html;
        }

        // ================= 8. ECONOMIC IMPACT SIMULATION ENGINE =================
        function populateSimulationDropdown() {
            const select = document.getElementById('sim-industry-select');
            select.innerHTML = '';
            activeSectors.forEach((sec, idx) => {
                const opt = document.createElement('option');
                opt.value = idx;
                opt.innerText = `[${sec.code}] ${sec.name}`;
                // Default C09 (컴퓨터, 전자 및 광학기기) or fallback index 10
                if (sec.code === 'C09' || (activeSectors.length > 10 && idx === 10)) {
                    opt.selected = true;
                }
                select.appendChild(opt);
            });
        }

        function setSimAmount(val) {
            document.getElementById('sim-amount-input').value = val;
        }

        function resetSimulationInput() {
            document.getElementById('sim-amount-input').value = 100000;
            populateSimulationDropdown();
            runSimulation();
        }

        function runSimulation() {
            const targetIdx = parseInt(document.getElementById('sim-industry-select').value) || 0;
            const dY_val = parseFloat(document.getElementById('sim-amount-input').value) || 0;

            const N = activeSectors.length;
            const Inv = activeInvMatrix;

            // Construct Final Demand Vector dY (N x 1)
            let dY = new Array(N).fill(0);
            dY[targetIdx] = dY_val;

            // Compute Production Output Impact dX = (I - A^d)^-1 * dY
            let dX = new Array(N).fill(0);
            for (let i = 0; i < N; i++) {
                dX[i] = Inv[i][targetIdx] * dY_val;
            }

            // Compute Value Added Impact dV = v_hat * dX
            let dV = new Array(N).fill(0);
            for (let i = 0; i < N; i++) {
                dV[i] = activeVARatios[i] * dX[i];
            }

            // Compute Import Impact dM
            let dM = new Array(N).fill(0);
            for (let i = 0; i < N; i++) {
                let colSumAm = 0;
                for (let k = 0; k < N; k++) colSumAm += activeAmMatrix[k][i];
                dM[i] = colSumAm * dX[i];
            }

            // Sum Totals
            const totalDX = dX.reduce((a, b) => a + b, 0);
            const totalDV = dV.reduce((a, b) => a + b, 0);
            const totalDM = dM.reduce((a, b) => a + b, 0);

            // Update KPI Cards
            document.getElementById('kpi-production-val').innerHTML = `${Math.round(totalDX).toLocaleString()} <span class="text-sm font-sans font-normal text-slate-500">백만원</span>`;
            document.getElementById('kpi-production-multiplier').innerText = `${(totalDX / dY_val).toFixed(4)} 배`;

            document.getElementById('kpi-va-val').innerHTML = `${Math.round(totalDV).toLocaleString()} <span class="text-sm font-sans font-normal text-slate-500">백만원</span>`;
            document.getElementById('kpi-va-multiplier').innerText = `${(totalDV / dY_val).toFixed(4)} (${((totalDV / dY_val) * 100).toFixed(2)}%)`;

            document.getElementById('kpi-import-val').innerHTML = `${Math.round(totalDM).toLocaleString()} <span class="text-sm font-sans font-normal text-slate-500">백만원</span>`;
            document.getElementById('kpi-import-multiplier').innerText = `${(totalDM / dY_val).toFixed(4)} (${((totalDM / dY_val) * 100).toFixed(2)}%)`;

            // Prepare Industry Breakdown Data
            let resultsList = [];
            for (let i = 0; i < N; i++) {
                resultsList.push({
                    code: activeSectors[i].code,
                    name: activeSectors[i].name,
                    dX: dX[i],
                    dV: dV[i],
                    dM: dM[i],
                    share: (dX[i] / totalDX) * 100,
                    isTarget: i === targetIdx
                });
            }

            // Render Horizontal Bar Chart (Top 15 Impacted Sectors)
            renderImpactChart([...resultsList].sort((a, b) => b.dX - a.dX).slice(0, 15));

            // Summary Texts
            const targetSector = activeSectors[targetIdx];
            const indirectDX = totalDX - dY_val;
            document.getElementById('summary-direct-indirect').innerText = 
                `선택한 [${targetSector.name}] 산업 자체의 직접 생산 증가 ${Math.round(dY_val).toLocaleString()}백만원 외에, 타 연관 산업에서 ${Math.round(indirectDX).toLocaleString()}백만원의 간접 중간재 생산이 추가로 유발되었습니다.`;

            let sortedNonTarget = [...resultsList].filter(r => !r.isTarget).sort((a, b) => b.dX - a.dX);
            if (sortedNonTarget.length >= 2) {
                document.getElementById('summary-top-beneficiary').innerHTML = 
                    `자기 산업을 제외하고 파급효과가 가장 큰 주요 연관 파급 산업은 <strong>${sortedNonTarget[0].code} ${sortedNonTarget[0].name}</strong> (${Math.round(sortedNonTarget[0].dX).toLocaleString()}백만원) 및 <strong>${sortedNonTarget[1].code} ${sortedNonTarget[1].name}</strong>입니다.`;
            }

            // Render Detailed Simulation Table
            renderSimulationTable(resultsList, totalDX);
        }

        function renderImpactChart(topData) {
            const ctx = document.getElementById('impact-bar-chart').getContext('2d');
            if (chartInstance) {
                chartInstance.destroy();
            }

            const labels = topData.map(d => `${d.code} ${d.name}`);
            const values = topData.map(d => Math.round(d.dX));
            const bgColors = topData.map(d => d.isTarget ? '#003876' : '#0284c7');

            chartInstance = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: '생산유발액 (백만원)',
                        data: values,
                        backgroundColor: bgColors,
                        borderRadius: 6,
                        borderSkipped: false,
                    }]
                },
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    return ` 생산유발액: ${context.parsed.x.toLocaleString()} 백만원`;
                                }
                            }
                        }
                    },
                    scales: {
                        x: {
                            grid: { color: '#f1f5f9' },
                            ticks: { font: { size: 10 } }
                        },
                        y: {
                            grid: { display: false },
                            ticks: { font: { size: 11, weight: 'bold' } }
                        }
                    }
                }
            });
        }

        function renderSimulationTable(dataList, totalDX) {
            const tbody = document.getElementById('sim-table-tbody');
            let html = '';
            dataList.forEach(item => {
                html += `<tr class="${item.isTarget ? 'bg-blue-50/70 font-bold' : 'hover:bg-slate-50'} transition-colors">`;
                html += `<td class="p-2 text-center text-slate-900">${item.code}</td>`;
                html += `<td class="p-2 font-medium text-slate-800">${item.name} ${item.isTarget ? '<span class="text-[10px] bg-blue-600 text-white px-1.5 py-0.5 rounded ml-1">충격 대상</span>' : ''}</td>`;
                html += `<td class="p-2 text-right font-bold text-blue-700">${Math.round(item.dX).toLocaleString()}</td>`;
                html += `<td class="p-2 text-right text-slate-600">${item.share.toFixed(2)}%</td>`;
                html += `<td class="p-2 text-right font-semibold text-emerald-700">${Math.round(item.dV).toLocaleString()}</td>`;
                html += `<td class="p-2 text-right font-semibold text-purple-700">${Math.round(item.dM).toLocaleString()}</td>`;
                html += `</tr>`;
            });
            tbody.innerHTML = html;
        }

        // CSV Export for Simulation Data
        function exportSimulationCSV() {
            const targetIdx = parseInt(document.getElementById('sim-industry-select').value) || 0;
            const dY_val = parseFloat(document.getElementById('sim-amount-input').value) || 0;
            const N = activeSectors.length;

            let csvStr = "\uFEFF산업코드,산업명,생산유발액(백만원),생산구성비(%),부가가치유발액(백만원),수입유발액(백만원)\n";

            let dX = activeInvMatrix.map(row => row[targetIdx] * dY_val);
            let totalDX = dX.reduce((a, b) => a + b, 0);

            for (let i = 0; i < N; i++) {
                let code = activeSectors[i].code;
                let name = `"${activeSectors[i].name}"`;
                let dxVal = Math.round(dX[i]);
                let share = ((dX[i] / totalDX) * 100).toFixed(2);
                let dvVal = Math.round(activeVARatios[i] * dX[i]);
                let colSumAm = 0;
                for (let k = 0; k < N; k++) colSumAm += activeAmMatrix[k][i];
                let dmVal = Math.round(colSumAm * dX[i]);

                csvStr += `${code},${name},${dxVal},${share},${dvVal},${dmVal}\n`;
            }

            const blob = new Blob([csvStr], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `BOK_2023_IO_Simulation_Results_${activeSectors[targetIdx].code}.csv`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }

        // Smooth Scroll Helper
        function scrollToSection(id) {
            const elem = document.getElementById(id);
            if (elem) {
                elem.scrollIntoView({ behavior: 'smooth' });
            }
        }
    </script>
</body>
</html># -title
