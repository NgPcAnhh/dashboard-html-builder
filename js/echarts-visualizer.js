/**
 * echarts-visualizer.js - Trình Biên Dịch JavaScript & Trực Quan Hóa Dữ Liệu Apache ECharts
 * Cung cấp môi trường live compiler chuyên sâu, thư viện mẫu biểu đồ đa dạng và liên kết file Excel/CSV.
 */

(function () {
  'use strict';

  // =========================================================================
  // 1. BIẾN TOÀN CỤC & DOM ELEMENTS
  // =========================================================================
  let myChart = null;
  let rawRows = [];
  let files = {};
  let currentFileName = "";
  let currentTheme = 'dark';

  // Elements
  const chartCanvas = document.getElementById('echartsMainCanvas');
  const codeEditor = document.getElementById('vizCodeEditor');
  const codeHighlight = document.getElementById('vizCodeHighlight');
  const codeHighlightContent = document.getElementById('vizCodeHighlightContent');
  const lineNumbers = document.getElementById('vizLineNumbers');
  const consoleBody = document.getElementById('vizConsoleBody');
  const btnRunCode = document.getElementById('btnRunCode');
  const btnResetCode = document.getElementById('btnResetCode');
  const btnCopyCode = document.getElementById('btnCopyCode');
  const btnFormatCode = document.getElementById('btnFormatCode');
  const btnExportPng = document.getElementById('btnExportPng');
  const btnOpenPresets = document.getElementById('btnOpenPresets');
  const btnOpenAiPrompt = document.getElementById('btnOpenAiPrompt');
  const btnClearConsole = document.getElementById('btnClearConsole');
  const themeToggleBtn = document.getElementById('vizThemeToggle');
  const renderTimeEl = document.getElementById('renderTime');
  const dataCountEl = document.getElementById('dataCount');

  // Tabs
  const tabBtns = document.querySelectorAll('.viz-tab-btn');
  const tabContents = document.querySelectorAll('.viz-tab-content');

  // Excel / CSV File Elements
  const dropZone = document.getElementById('vizDropZone');
  const fileInput = document.getElementById('vizFileInput');
  const dataMetaEl = document.getElementById('vizDataMeta');
  const previewTableContainer = document.getElementById('vizPreviewTableContainer');
  const btnAutoPlotData = document.getElementById('btnAutoPlotData');

  // Dữ liệu mẫu dự phòng khi chưa đẩy file (để biến 'data' luôn có sẵn)
  const DEFAULT_SAMPLE_DATA = [
    { "Khu Vực": "Hà Nội", "Doanh Thu": 450000000, "Đơn Hàng": 120, "Lợi Nhuận": 135000000 },
    { "Khu Vực": "TP.HCM", "Doanh Thu": 720000000, "Đơn Hàng": 195, "Lợi Nhuận": 216000000 },
    { "Khu Vực": "Đà Nẵng", "Doanh Thu": 280000000, "Đơn Hàng": 84, "Lợi Nhuận": 84000000 },
    { "Khu Vực": "Hải Phòng", "Doanh Thu": 310000000, "Đơn Hàng": 92, "Lợi Nhuận": 93000000 },
    { "Khu Vực": "Cần Thơ", "Doanh Thu": 190000000, "Đơn Hàng": 60, "Lợi Nhuận": 57000000 },
    { "Khu Vực": "Bình Dương", "Doanh Thu": 380000000, "Đơn Hàng": 110, "Lợi Nhuận": 114000000 }
  ];

  // Modals
  const presetModal = document.getElementById('vizPresetModal');
  const btnClosePresetModal = document.getElementById('btnClosePresetModal');
  const aiPromptModal = document.getElementById('vizAiPromptModal');
  const btnCloseAiModal = document.getElementById('btnCloseAiModal');
  const btnGenerateAiPrompt = document.getElementById('btnGenerateAiPrompt');
  const btnCopyAiPrompt = document.getElementById('btnCopyAiPrompt');
  const aiUserRequirement = document.getElementById('aiUserRequirement');
  const aiPromptOutput = document.getElementById('aiPromptOutput');

  // =========================================================================
  // 2. KHO MẪU BIỂU ĐỒ ECHARTS PHONG PHÚ (CHART PRESETS)
  // =========================================================================
  const CHART_PRESETS = [
    {
      id: 'dashboard-28-col',
      name: 'Dashboard Phân Tích 28 Cột (Interactive)',
      icon: '🎛️',
      category: 'Phân Tích Nâng Cao',
      desc: 'Dashboard tương tác trực tiếp: Nút bấm chuyển đổi Khu Vực / Danh Mục / Kênh Bán / Thanh Toán, chuyển chỉ số Doanh Thu / Lợi Nhuận / Số Lượng ngay trong chart.',
      code: `// =========================================================================
// PHÂN TÍCH ĐA CHIỀU FILE 28 CỘT (DOANH THU, LỢI NHUẬN, KÊNH BÁN, KHU VỰC)
// Hỗ trợ chuyển đổi chiều phân tích và chỉ số trực tiếp bằng nút bấm trong chart!
// =========================================================================

const rows = (typeof data !== 'undefined' && data && data.length > 0) ? data : [];

// 1. Khởi tạo trạng thái mặc định
let currentDim = 'Khu vực';       // 'Khu vực' | 'Danh mục' | 'Kênh bán' | 'Phương thức TT'
let currentMetric = 'finance';    // 'finance' (Doanh thu & Lợi nhuận) | 'quantity' (Số lượng) | 'pie' (Tỷ trọng Doanh thu)

// 2. Hàm tổng hợp dữ liệu theo chiều đã chọn
function aggregate(dim) {
  const map = {};
  rows.forEach(r => {
    const key = String(r[dim] || 'Chưa phân loại').trim();
    if (!map[key]) {
      map[key] = { revenue: 0, cost: 0, profit: 0, qty: 0, orders: 0 };
    }
    const rev = typeof r['Doanh thu'] === 'number' ? r['Doanh thu'] : parseFloat(String(r['Doanh thu'] || '').replace(/[^0-9.-]/g, '')) || 0;
    const pro = typeof r['Lợi nhuận'] === 'number' ? r['Lợi nhuận'] : parseFloat(String(r['Lợi nhuận'] || '').replace(/[^0-9.-]/g, '')) || 0;
    const q = typeof r['Số lượng'] === 'number' ? r['Số lượng'] : parseFloat(String(r['Số lượng'] || '').replace(/[^0-9.-]/g, '')) || 0;

    map[key].revenue += rev;
    map[key].profit += pro;
    map[key].qty += q;
    map[key].orders += 1;
  });

  // Sắp xếp giảm dần theo doanh thu
  const sortedKeys = Object.keys(map).sort((a, b) => map[b].revenue - map[a].revenue);
  return {
    categories: sortedKeys,
    revenues: sortedKeys.map(k => Math.round(map[k].revenue)),
    profits: sortedKeys.map(k => Math.round(map[k].profit)),
    quantities: sortedKeys.map(k => map[k].qty),
    orders: sortedKeys.map(k => map[k].orders),
    pieData: sortedKeys.map(k => ({ name: k, value: Math.round(map[k].revenue) }))
  };
}

// 3. Hàm tạo option ECharts động theo chiều và chỉ số
function getDashboardOption(dim, metric) {
  const agg = aggregate(dim);
  const totalRev = agg.revenues.reduce((a, b) => a + b, 0);
  const totalProfit = agg.profits.reduce((a, b) => a + b, 0);

  // Tạo các nút bấm tương tác Graphic bên trong Chart
  const dimsList = [
    { key: 'Khu vực', label: '🏢 Khu Vực' },
    { key: 'Danh mục', label: '📦 Danh Mục' },
    { key: 'Kênh bán', label: '🛒 Kênh Bán' },
    { key: 'Phương thức TT', label: '💳 Thanh Toán' }
  ];

  const metricsList = [
    { key: 'finance', label: '💰 Doanh Thu & Lợi Nhuận' },
    { key: 'quantity', label: '📦 Số Lượng & Đơn Hàng' },
    { key: 'pie', label: '🍩 Cơ Cấu Tỷ Trọng' }
  ];

  const graphicElements = [];

  // Group 1: Nút chọn Chiều phân tích (Dimension Tabs)
  dimsList.forEach((d, idx) => {
    const isAct = d.key === dim;
    graphicElements.push({
      type: 'group',
      left: 20 + idx * 115,
      top: 50,
      cursor: 'pointer',
      onclick: function() {
        currentDim = d.key;
        myChart.setOption(getDashboardOption(currentDim, currentMetric), true);
      },
      children: [
        {
          type: 'rect',
          shape: { width: 105, height: 26, r: 4 },
          style: {
            fill: isAct ? '#2563eb' : '#f1f5f9',
            stroke: isAct ? '#1d4ed8' : '#cbd5e1',
            lineWidth: 1
          }
        },
        {
          type: 'text',
          style: {
            text: d.label,
            x: 52,
            y: 13,
            textAlign: 'center',
            textVerticalAlign: 'middle',
            fill: isAct ? '#ffffff' : '#334155',
            font: isAct ? 'bold 11px sans-serif' : '11px sans-serif'
          }
        }
      ]
    });
  });

  // Group 2: Nút chọn Chỉ số (Metric Tabs)
  metricsList.forEach((m, idx) => {
    const isAct = m.key === metric;
    graphicElements.push({
      type: 'group',
      right: 20 + (metricsList.length - 1 - idx) * 165,
      top: 50,
      cursor: 'pointer',
      onclick: function() {
        currentMetric = m.key;
        myChart.setOption(getDashboardOption(currentDim, currentMetric), true);
      },
      children: [
        {
          type: 'rect',
          shape: { width: 155, height: 26, r: 4 },
          style: {
            fill: isAct ? '#059669' : '#f1f5f9',
            stroke: isAct ? '#047857' : '#cbd5e1',
            lineWidth: 1
          }
        },
        {
          type: 'text',
          style: {
            text: m.label,
            x: 77,
            y: 13,
            textAlign: 'center',
            textVerticalAlign: 'middle',
            fill: isAct ? '#ffffff' : '#334155',
            font: isAct ? 'bold 11px sans-serif' : '11px sans-serif'
          }
        }
      ]
    });
  });

  // Cấu hình series theo loại chỉ số
  let seriesList = [];
  let yAxisConfig = [];
  let isPie = metric === 'pie';

  if (metric === 'finance') {
    yAxisConfig = [
      {
        type: 'value',
        name: 'Doanh Thu (VNĐ)',
        axisLabel: {
          formatter: v => v >= 1e9 ? (v / 1e9).toFixed(1) + ' Tỷ' : (v >= 1e6 ? (v / 1e6).toFixed(0) + ' Tr' : v)
        },
        splitLine: { lineStyle: { type: 'dashed', opacity: 0.25 } }
      },
      {
        type: 'value',
        name: 'Tỷ Suất LN (%)',
        min: 0,
        max: 60,
        axisLabel: { formatter: '{value}%' },
        splitLine: { show: false }
      }
    ];

    const margins = agg.revenues.map((rev, i) => rev > 0 ? Number(((agg.profits[i] / rev) * 100).toFixed(1)) : 0);

    seriesList = [
      {
        name: 'Doanh Thu',
        type: 'bar',
        barMaxWidth: 35,
        itemStyle: {
          borderRadius: [6, 6, 0, 0],
          color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: '#38bdf8' },
            { offset: 1, color: '#2563eb' }
          ])
        },
        data: agg.revenues,
        markPoint: { data: [{ type: 'max', name: 'Cao nhất' }] }
      },
      {
        name: 'Lợi Nhuận',
        type: 'bar',
        barMaxWidth: 35,
        itemStyle: {
          borderRadius: [6, 6, 0, 0],
          color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: '#34d399' },
            { offset: 1, color: '#059669' }
          ])
        },
        data: agg.profits
      },
      {
        name: 'Tỷ Suất Lợi Nhuận (%)',
        type: 'line',
        yAxisIndex: 1,
        smooth: true,
        lineStyle: { width: 3, color: '#f59e0b' },
        itemStyle: { color: '#f59e0b', borderWidth: 2, borderColor: '#ffffff' },
        data: margins
      }
    ];
  } else if (metric === 'quantity') {
    yAxisConfig = [
      {
        type: 'value',
        name: 'Số Lượng Sản Phẩm',
        splitLine: { lineStyle: { type: 'dashed', opacity: 0.25 } }
      },
      {
        type: 'value',
        name: 'Số Đơn Hàng',
        splitLine: { show: false }
      }
    ];

    seriesList = [
      {
        name: 'Số Lượng SP',
        type: 'bar',
        barMaxWidth: 35,
        itemStyle: {
          borderRadius: [6, 6, 0, 0],
          color: '#6366f1'
        },
        data: agg.quantities
      },
      {
        name: 'Số Đơn Hàng',
        type: 'line',
        yAxisIndex: 1,
        smooth: true,
        lineStyle: { width: 3, color: '#ec4899' },
        itemStyle: { color: '#ec4899', borderWidth: 2, borderColor: '#ffffff' },
        data: agg.orders
      }
    ];
  } else {
    // Pie / Donut
    seriesList = [
      {
        name: 'Doanh Thu',
        type: 'pie',
        radius: ['40%', '70%'],
        center: ['50%', '58%'],
        avoidLabelOverlap: false,
        itemStyle: {
          borderRadius: 8,
          borderColor: '#ffffff',
          borderWidth: 2
        },
        label: {
          show: true,
          formatter: '{b}: {d}%'
        },
        emphasis: {
          label: { show: true, fontSize: 14, fontWeight: 'bold' }
        },
        data: agg.pieData
      }
    ];
  }

  return {
    backgroundColor: '#ffffff',
    title: {
      text: \`BÁO CÁO PHÂN TÍCH THEO \${dim.toUpperCase()}\`,
      subtext: \`Tổng Doanh Thu: \${totalRev.toLocaleString('vi-VN')} đ | Lợi Nhuận: \${totalProfit.toLocaleString('vi-VN')} đ | Bấm nút phía trên để đổi góc nhìn & chỉ số\`,
      left: 'center',
      top: 10,
      textStyle: { fontSize: 15, fontWeight: 'bold', color: '#0f172a' },
      subtextStyle: { fontSize: 11.5, color: '#475569' }
    },
    tooltip: {
      trigger: isPie ? 'item' : 'axis',
      axisPointer: { type: 'shadow' },
      formatter: function(params) {
        if (isPie) {
          return \`<b>\${params.name}</b><br/>Doanh Thu: <b>\${params.value.toLocaleString('vi-VN')} đ</b> (\${params.percent}%)\`;
        }
        let tip = \`<b>\${params[0].name}</b><br/>\`;
        params.forEach(p => {
          let valStr = p.value.toLocaleString('vi-VN');
          if (p.seriesName.includes('%')) valStr += ' %';
          else if (p.seriesName.includes('Doanh') || p.seriesName.includes('Lợi')) valStr += ' đ';
          tip += \`\${p.marker} \${p.seriesName}: <b>\${valStr}</b><br/>\`;
        });
        return tip;
      }
    },
    legend: {
      show: true,
      top: 85,
      textStyle: { color: '#334155' }
    },
    toolbox: {
      feature: {
        saveAsImage: { title: 'Tải ảnh PNG' },
        magicType: isPie ? {} : { type: ['bar', 'line', 'stack'], title: { bar: 'Cột', line: 'Đường', stack: 'Chồng' } },
        dataView: { title: 'Xem số liệu', readOnly: true },
        restore: { title: 'Khôi phục' }
      },
      top: 85,
      right: 25
    },
    grid: isPie ? undefined : {
      left: '4%', right: '4%', bottom: '12%', top: 125,
      containLabel: true
    },
    dataZoom: isPie ? undefined : [
      { type: 'inside', start: 0, end: 100 },
      { type: 'slider', start: 0, end: 100, bottom: 8 }
    ],
    xAxis: isPie ? undefined : {
      type: 'category',
      data: agg.categories,
      axisLabel: { interval: 0, rotate: agg.categories.length > 6 ? 20 : 0, color: '#334155' }
    },
    yAxis: isPie ? undefined : yAxisConfig,
    graphic: graphicElements,
    series: seriesList
  };
}

// 4. Khởi chạy và thiết lập biểu đồ
const option = getDashboardOption(currentDim, currentMetric);
return option;`
    },
    {
      id: 'bar-gradient',
      name: 'Cột Doanh Thu & Lợi Nhuận',
      icon: '📊',
      category: 'So Sánh',
      desc: 'Biểu đồ cột đôi dải màu Gradient, nhãn giá trị động và đường trung bình (markLine).',
      code: `// 1. Dữ liệu các tháng kinh doanh
const months = ['Tháng 1', 'Tháng 2', 'Tháng 3', 'Tháng 4', 'Tháng 5', 'Tháng 6', 'Tháng 7', 'Tháng 8', 'Tháng 9', 'Tháng 10', 'Tháng 11', 'Tháng 12'];
const revenue = [420, 532, 601, 734, 890, 1030, 920, 1150, 1280, 1420, 1350, 1680]; // Triệu VNĐ
const profit = [120, 160, 190, 240, 310, 380, 320, 410, 470, 530, 490, 620];

// 2. Cấu hình Apache ECharts
const option = {
  title: {
    text: 'Báo Cáo Doanh Thu & Lợi Nhuận 2026',
    subtext: 'Đơn vị tính: Triệu VNĐ (Dữ liệu mẫu)',
    left: 'center',
    textStyle: { fontWeight: 'bold' }
  },
  tooltip: {
    trigger: 'axis',
    axisPointer: { type: 'shadow' }
  },
  legend: {
    top: '38px',
    data: ['Doanh Thu', 'Lợi Nhuận']
  },
  grid: {
    left: '3%',
    right: '4%',
    bottom: '8%',
    top: '80px',
    containLabel: true
  },
  toolbox: {
    feature: {
      saveAsImage: { title: 'Tải ảnh' },
      dataView: { title: 'Xem dữ liệu', readOnly: true },
      magicType: { type: ['bar', 'line'], title: { bar: 'Cột', line: 'Đường' } }
    }
  },
  xAxis: {
    type: 'category',
    data: months,
    axisLabel: { interval: 0, rotate: 25 }
  },
  yAxis: {
    type: 'value',
    name: 'Triệu VNĐ',
    splitLine: { lineStyle: { type: 'dashed', opacity: 0.25 } }
  },
  series: [
    {
      name: 'Doanh Thu',
      type: 'bar',
      barGap: '15%',
      barMaxWidth: 24,
      itemStyle: {
        borderRadius: [6, 6, 0, 0],
        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
          { offset: 0, color: '#38bdf8' },
          { offset: 1, color: '#2563eb' }
        ])
      },
      data: revenue,
      markPoint: {
        data: [
          { type: 'max', name: 'Đỉnh' },
          { type: 'min', name: 'Đáy' }
        ]
      }
    },
    {
      name: 'Lợi Nhuận',
      type: 'bar',
      barMaxWidth: 24,
      itemStyle: {
        borderRadius: [6, 6, 0, 0],
        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
          { offset: 0, color: '#34d399' },
          { offset: 1, color: '#059669' }
        ])
      },
      data: profit,
      markLine: {
        data: [{ type: 'average', name: 'TB' }]
      }
    }
  ]
};

return option;`
    },
    {
      id: 'area-smooth',
      name: 'Đường Xu Hướng Vùng (Area Smooth)',
      icon: '📈',
      category: 'Xu Hướng',
      desc: 'Biểu đồ đường mượt mà kèm vùng phủ gradient mềm mại, zoom dữ liệu tương tác.',
      code: `// Chuỗi ngày trong tháng
const dates = Array.from({ length: 30 }, (_, i) => \`Ngày \${i + 1}\`);
const onlineUsers = [
  320, 380, 420, 510, 680, 720, 650, 780, 890, 950,
  1100, 1250, 1180, 1340, 1420, 1580, 1720, 1650, 1820, 1950,
  2100, 2300, 2250, 2480, 2600, 2750, 2900, 3100, 3050, 3400
];

const option = {
  title: {
    text: 'Xu Hướng Tăng Trưởng Người Dùng Hoạt Động (DAU)',
    subtext: 'Theo dõi 30 ngày qua với bộ điều khiển DataZoom',
    left: 'center'
  },
  tooltip: {
    trigger: 'axis',
    formatter: '{b}<br/>👥 Người dùng: <b>{c}</b>'
  },
  grid: {
    left: '4%',
    right: '4%',
    bottom: '14%',
    top: '75px',
    containLabel: true
  },
  dataZoom: [
    { type: 'inside', start: 30, end: 100 },
    { type: 'slider', start: 30, end: 100 }
  ],
  xAxis: {
    type: 'category',
    boundaryGap: false,
    data: dates
  },
  yAxis: {
    type: 'value',
    splitLine: { lineStyle: { type: 'dashed', opacity: 0.25 } }
  },
  series: [
    {
      name: 'Người dùng',
      type: 'line',
      smooth: true,
      symbol: 'circle',
      symbolSize: 6,
      lineStyle: { width: 3, color: '#38bdf8' },
      itemStyle: { color: '#38bdf8', borderWidth: 2, borderColor: '#fff' },
      areaStyle: {
        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
          { offset: 0, color: 'rgba(56, 189, 248, 0.45)' },
          { offset: 1, color: 'rgba(37, 99, 235, 0.02)' }
        ])
      },
      data: onlineUsers,
      markPoint: {
        data: [{ type: 'max', name: 'Kỷ lục' }]
      }
    }
  ]
};

return option;`
    },
    {
      id: 'pie-donut',
      name: 'Cơ Cấu Tỷ Trọng (Donut Chart)',
      icon: '🍩',
      category: 'Cơ Cấu',
      desc: 'Biểu đồ bánh Donut với góc bo tròn mềm mại, chú giải bên hông và tooltip chi tiết.',
      code: `const productShare = [
  { value: 1450, name: 'Tín Dụng & Cho Vay' },
  { value: 1120, name: 'Thẻ & Ngân Hàng Số' },
  { value: 890, name: 'Tiền Gửi Tiết Kiệm' },
  { value: 640, name: 'Bảo Hiểm & Đầu Tư' },
  { value: 430, name: 'Dịch Vụ Thanh Toán' },
  { value: 270, name: 'Kinh Doanh Ngoại Tệ' }
];

const option = {
  title: {
    text: 'Cơ Cấu Doanh Thu Theo Khối Sản Phẩm',
    subtext: 'Tỷ trọng đóng góp lũy kế (Tỷ VNĐ)',
    left: 'center'
  },
  tooltip: {
    trigger: 'item',
    formatter: '{a} <br/><b>{b}</b>: {c} Tỷ ({d}%)'
  },
  legend: {
    orient: 'vertical',
    left: '5%',
    top: 'center'
  },
  series: [
    {
      name: 'Khối Sản Phẩm',
      type: 'pie',
      radius: ['45%', '72%'],
      center: ['60%', '55%'],
      avoidLabelOverlap: false,
      itemStyle: {
        borderRadius: 8,
        borderColor: '#ffffff',
        borderWidth: 2
      },
      label: {
        show: false,
        position: 'center'
      },
      emphasis: {
        label: {
          show: true,
          fontSize: 16,
          fontWeight: 'bold',
          formatter: '{b}\\n{d}%'
        },
        itemStyle: {
          shadowBlur: 10,
          shadowOffsetX: 0,
          shadowColor: 'rgba(0, 0, 0, 0.5)'
        }
      },
      labelLine: { show: false },
      data: productShare
    }
  ]
};

return option;`
    },
    {
      id: 'mixed-combo',
      name: 'Cột & Đường Hỗn Hợp (Hai Trục Y)',
      icon: '📶',
      category: 'Phân Tích Đa Chiều',
      desc: 'Cột thể hiện Doanh số (trục Y trái), Đường thể hiện Tỷ lệ tăng trưởng % (trục Y phải).',
      code: `const months = ['T1', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'T8', 'T9', 'T10', 'T11', 'T12'];
const sales = [52, 68, 79, 85, 110, 135, 128, 142, 160, 175, 190, 220]; // Tỷ
const growthRates = [5.2, 8.4, 12.1, 7.6, 29.4, 22.7, -5.2, 10.9, 12.7, 9.4, 8.6, 15.8]; // %

const option = {
  title: {
    text: 'Doanh Số Bán Hàng vs Tốc Độ Tăng Trưởng',
    left: 'center'
  },
  tooltip: {
    trigger: 'axis',
    axisPointer: { type: 'cross' }
  },
  legend: {
    top: '35px',
    data: ['Doanh Số', 'Tăng Trưởng %']
  },
  grid: {
    left: '4%', right: '4%', bottom: '8%', top: '80px',
    containLabel: true
  },
  xAxis: [
    { type: 'category', data: months, axisPointer: { type: 'shadow' } }
  ],
  yAxis: [
    {
      type: 'value',
      name: 'Doanh số (Tỷ)',
      min: 0,
      max: 250,
      splitLine: { lineStyle: { type: 'dashed', opacity: 0.2 } }
    },
    {
      type: 'value',
      name: 'Tăng trưởng (%)',
      min: -15,
      max: 40,
      axisLabel: { formatter: '{value} %' },
      splitLine: { show: false }
    }
  ],
  series: [
    {
      name: 'Doanh Số',
      type: 'bar',
      barMaxWidth: 26,
      itemStyle: {
        borderRadius: [4, 4, 0, 0],
        color: '#2563eb'
      },
      data: sales
    },
    {
      name: 'Tăng Trưởng %',
      type: 'line',
      yAxisIndex: 1,
      smooth: true,
      lineStyle: { width: 3, color: '#f59e0b' },
      itemStyle: { color: '#f59e0b' },
      data: growthRates
    }
  ]
};

return option;`
    },
    {
      id: 'radar-kpi',
      name: 'Radar Đánh Giá Năng Lực & KPI',
      icon: '🎯',
      category: 'Đánh Giá',
      desc: 'Biểu đồ mạng nhện so sánh đa chỉ số năng lực giữa Kế hoạch và Thực tế đạt được.',
      code: `const option = {
  title: {
    text: 'Đánh Giá Năng Lực & Hiệu Suất Hoạt Động (KPI)',
    subtext: 'So sánh Thực tế vs Mục tiêu giao khoán',
    left: 'center'
  },
  tooltip: {},
  legend: {
    top: '35px',
    data: ['Mục Tiêu Chỉ Tiêu', 'Thực Tế Đạt Được']
  },
  radar: {
    center: ['50%', '55%'],
    radius: '65%',
    axisName: {
      color: '#334155',
      fontWeight: '600'
    },
    indicator: [
      { name: 'Tăng Trưởng Doanh Số', max: 100 },
      { name: 'Khách Hàng Mới', max: 100 },
      { name: 'Chất Lượng Dịch Vụ (CSAT)', max: 100 },
      { name: 'Kiểm Soát Rủi Ro', max: 100 },
      { name: 'Năng Suất Nhân Sự', max: 100 },
      { name: 'Chuyển Đổi Số', max: 100 }
    ],
    splitArea: {
      show: true,
      areaStyle: {
        color: ['#f8fafc', '#f1f5f9']
      }
    }
  },
  series: [
    {
      name: 'KPI Performance',
      type: 'radar',
      data: [
        {
          value: [85, 90, 80, 85, 75, 90],
          name: 'Mục Tiêu Chỉ Tiêu',
          itemStyle: { color: '#60a5fa' },
          areaStyle: { color: 'rgba(96, 165, 250, 0.25)' }
        },
        {
          value: [95, 82, 92, 88, 85, 96],
          name: 'Thực Tế Đạt Được',
          itemStyle: { color: '#10b981' },
          areaStyle: { color: 'rgba(16, 185, 129, 0.35)' }
        }
      ]
    }
  ]
};

return option;`
    },
    {
      id: 'funnel-sales',
      name: 'Phễu Chuyển Đổi Khách Hàng (Funnel)',
      icon: '⏳',
      category: 'Phễu / Quy Trình',
      desc: 'Phễu chuyển đổi bán hàng từ Tiếp cận -> Tương tác -> Báo giá -> Hợp đồng thành công.',
      code: `const funnelData = [
  { value: 100, name: 'Lượt Tiếp Cận (Reach)', realCount: '150,000' },
  { value: 72, name: 'Quan Tâm & Xem Sản Phẩm', realCount: '108,000' },
  { value: 48, name: 'Tạo Tài Khoản & Điền Form', realCount: '72,000' },
  { value: 28, name: 'Được Tư Vấn & Nhận Báo Giá', realCount: '42,000' },
  { value: 14, name: 'Chốt Hợp Đồng / Kích Hoạt', realCount: '21,000' }
];

const option = {
  title: {
    text: 'Phễu Chuyển Đổi Hành Trình Khách Hàng',
    subtext: 'Tỷ lệ % chuyển đổi lũy kế qua từng giai đoạn',
    left: 'center'
  },
  tooltip: {
    trigger: 'item',
    formatter: function(params) {
      return \`<b>\${params.data.name}</b><br/>Tỷ lệ: <b>\${params.data.value}%</b><br/>Số lượng: <b>\${params.data.realCount}</b> khách\`;
    }
  },
  legend: {
    top: '35px',
    data: funnelData.map(d => d.name)
  },
  series: [
    {
      name: 'Hành Trình',
      type: 'funnel',
      left: '15%',
      top: 80,
      bottom: 30,
      width: '70%',
      min: 0,
      max: 100,
      minSize: '15%',
      maxSize: '100%',
      sort: 'descending',
      gap: 3,
      label: {
        show: true,
        position: 'inside',
        formatter: '{b}: {c}%',
        color: '#fff',
        fontWeight: 'bold'
      },
      itemStyle: {
        borderColor: '#ffffff',
        borderWidth: 2
      },
      data: funnelData
    }
  ]
};

return option;`
    },
    {
      id: 'gauge-progress',
      name: 'Đồng Hồ Đo Hiệu Suất (Gauge Chart)',
      icon: '🧭',
      category: 'Chỉ Số',
      desc: 'Đồng hồ tốc độ / tỷ lệ hoàn thành kế hoạch với dải màu an toàn - cảnh báo - xuất sắc.',
      code: `const completionRate = 86.4; // Tỷ lệ hoàn thành %

const option = {
  title: {
    text: 'Tiến Độ Hoàn Thành Kế Hoạch Năm 2026',
    subtext: 'Chỉ số đo lường hiệu năng tổng hợp (Score)',
    left: 'center'
  },
  series: [
    {
      type: 'gauge',
      center: ['50%', '60%'],
      radius: '85%',
      startAngle: 200,
      endAngle: -20,
      min: 0,
      max: 100,
      splitNumber: 10,
      axisLine: {
        lineStyle: {
          width: 18,
          color: [
            [0.6, '#ef4444'],   // Dưới 60%: Đỏ cảnh báo
            [0.85, '#f59e0b'],  // 60% - 85%: Vàng khá
            [1, '#10b981']     // Trên 85%: Xanh xuất sắc
          ]
        }
      },
      pointer: {
        itemStyle: { color: '#38bdf8' },
        width: 6,
        length: '65%'
      },
      axisTick: { distance: -18, length: 6, lineStyle: { color: '#cbd5e1', width: 1 } },
      splitLine: { distance: -18, length: 14, lineStyle: { color: '#94a3b8', width: 2 } },
      axisLabel: { color: '#475569', distance: 25, fontSize: 11 },
      detail: {
        valueAnimation: true,
        formatter: '{value}%',
        color: '#0f172a',
        fontSize: 28,
        fontWeight: 'bold',
        offsetCenter: [0, '40%']
      },
      data: [{ value: completionRate, name: 'ĐÃ ĐẠT ĐƯỢC' }],
      title: { offsetCenter: [0, '65%'], fontSize: 13, color: '#475569' }
    }
  ]
};

return option;`
    },
    {
      id: 'excel-dynamic',
      name: 'Vẽ Từ Dữ Liệu Excel / CSV Đã Nạp',
      icon: '⚡',
      category: 'Dữ Liệu Thực Tế',
      desc: 'Tự động đọc các dòng từ biến rawRows của file Excel để thống kê và vẽ biểu đồ động.',
      code: `// Biến 'rawRows' chứa toàn bộ dòng dữ liệu từ file Excel/CSV bạn nạp vào tab Dữ liệu.
// Nếu chưa nạp file, hệ thống sẽ tự động dùng dữ liệu giả lập bên dưới:
let data = (typeof rawRows !== 'undefined' && rawRows && rawRows.length > 0)
  ? rawRows
  : [
      { "Khu Vực": "Hà Nội", "Doanh Thu": 450000000, "Đơn Hàng": 120 },
      { "Khu Vực": "TP.HCM", "Doanh Thu": 720000000, "Đơn Hàng": 195 },
      { "Khu Vực": "Đà Nẵng", "Doanh Thu": 280000000, "Đơn Hàng": 84 },
      { "Khu Vực": "Hải Phòng", "Doanh Thu": 310000000, "Đơn Hàng": 92 },
      { "Khu Vực": "Cần Thơ", "Doanh Thu": 190000000, "Đơn Hàng": 60 },
      { "Khu Vực": "Bình Dương", "Doanh Thu": 380000000, "Đơn Hàng": 110 }
    ];

// Tự động tìm cột nhãn (Text) và cột số (Number)
const firstRow = data[0] || {};
const keys = Object.keys(firstRow);
const labelCol = keys[0] || 'Khu Vực';
const valueCol = keys[1] || 'Doanh Thu';

// Gom nhóm hoặc lấy danh sách
const categories = data.map(r => String(r[labelCol] || ''));
const values = data.map(r => {
  const v = r[valueCol];
  return typeof v === 'number' ? v : parseFloat(String(v || '').replace(/[^0-9.-]/g, '')) || 0;
});

const option = {
  title: {
    text: \`Thống Kê Dữ Liệu: \${valueCol} Theo \${labelCol}\`,
    subtext: \`Tổng cộng \${data.length.toLocaleString('vi-VN')} dòng dữ liệu từ Excel\`,
    left: 'center'
  },
  tooltip: {
    trigger: 'axis',
    axisPointer: { type: 'shadow' }
  },
  grid: {
    left: '3%', right: '4%', bottom: '10%', top: '75px',
    containLabel: true
  },
  xAxis: {
    type: 'category',
    data: categories,
    axisLabel: { interval: 0, rotate: 20 }
  },
  yAxis: {
    type: 'value',
    name: valueCol,
    splitLine: { lineStyle: { type: 'dashed', opacity: 0.25 } }
  },
  series: [
    {
      name: valueCol,
      type: 'bar',
      barMaxWidth: 30,
      itemStyle: {
        borderRadius: [6, 6, 0, 0],
        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
          { offset: 0, color: '#38bdf8' },
          { offset: 1, color: '#1d4ed8' }
        ])
      },
      data: values,
      markPoint: { data: [{ type: 'max', name: 'Cao nhất' }] },
      markLine: { data: [{ type: 'average', name: 'Trung bình' }] }
    }
  ]
};

return option;`
    }
  ];

  // =========================================================================
  // 3. KHỞI TẠO ECHARTS VÀ LINE NUMBERS (NỀN TRẮNG MẶC ĐỊNH)
  // =========================================================================
  function initChart() {
    if (!chartCanvas) return;
    if (myChart) {
      myChart.dispose();
    }
    
    // Nền của chart hiển thị mặc định luôn là màu trắng (null theme chuẩn)
    // Nếu trong code ECharts người dùng muốn màu khác thì cấu hình qua backgroundColor
    myChart = echarts.init(chartCanvas, null, {
      renderer: 'canvas'
    });

    window.addEventListener('resize', () => {
      if (myChart) myChart.resize();
    });
  }

  // =========================================================================
  // 3. SYNTAX HIGHLIGHTING & AUTO FORMATTER (TÔ MÀU & CĂN CHỈNH CODE TỰ ĐỘNG)
  // =========================================================================
  function escapeHtml(str) {
    return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  /**
   * Tạo thẻ HTML cho đường dóng và dải màu thụt dòng (Indent Rainbow Guides)
   * Mỗi 2 khoảng trắng (hoặc 1 tab) tương ứng 1 cấp độ thụt dòng
   */
  function renderIndentGuides(indentStr) {
    if (!indentStr) return '';
    let html = '';
    let level = 1;
    let i = 0;

    while (i < indentStr.length) {
      const ch = indentStr[i];
      const lvlClass = `indent-lvl-${((level - 1) % 6) + 1}`;

      if (ch === '\t') {
        html += `<span class="indent-guide ${lvlClass}">\t</span>`;
        level++;
        i++;
      } else if (ch === ' ' && indentStr[i + 1] === ' ') {
        html += `<span class="indent-guide ${lvlClass}">  </span>`;
        level++;
        i += 2;
      } else {
        // Khoảng trắng lẻ (1 space)
        html += `<span class="indent-guide ${lvlClass}">${ch}</span>`;
        i++;
      }
    }

    return html;
  }

  // Regex nhận diện các token JavaScript cho cấu hình ECharts chuẩn xác
  const JS_TOKEN_REGEX = /(\/\*[\s\S]*?\*\/|\/\/[^\n]*)|("(?:\\.|[^"\n\\])*"|'(?:\\.|[^'\n\\])*'|`(?:\\.|[^`\\])*`|`(?:\\.|[^`\\])*$)|(\b(?:const|let|var|function|return|if|else|for|while|do|switch|case|default|break|continue|try|catch|finally|throw|new|typeof|instanceof|class|extends|import|export|from|async|await|yield|in|of|void|delete|debugger)\b)|(\b(?:true|false|null|undefined|NaN|Infinity)\b)|(\b(?:echarts|myChart|option|data|rawRows|files|fileSheets|getFile|getSheet|mergeAllRows|context|dataColumns|sampleRecords|console|window|document|Array|Object|String|Number|Boolean|Math|JSON|Date|RegExp|Map|Set|Promise|XLSX)\b)|(\b\d+(?:\.\d+)?(?:e[+-]?\d+)?\b)|(\b[a-zA-Z_$][a-zA-Z0-9_$]*(?=\s*\())|(\b[a-zA-Z_$][a-zA-Z0-9_$]*(?=\s*:))|([{}()\[\]])|(===|!==|==|!=|<=|>=|=>|&&|\|\||\?\?|\+=|-=|\*=|\/=|[\+\-\*\/%=\?:\!&\|\^~<>])/g;

  /**
   * Tô màu cú pháp từng dòng code JavaScript & ECharts
   */
  function highlightCodeLine(codePart, state) {
    if (state.inBlockComment) {
      const closeIdx = codePart.indexOf('*/');
      if (closeIdx !== -1) {
        const commentText = codePart.substring(0, closeIdx + 2);
        state.inBlockComment = false;
        return `<span class="token-comment">${escapeHtml(commentText)}</span>` + highlightCodeLine(codePart.substring(closeIdx + 2), state);
      } else {
        return `<span class="token-comment">${escapeHtml(codePart)}</span>`;
      }
    }

    if (state.inTemplateStr) {
      let closeIdx = -1;
      for (let i = 0; i < codePart.length; i++) {
        if (codePart[i] === '`' && (i === 0 || codePart[i - 1] !== '\\')) {
          closeIdx = i;
          break;
        }
      }
      if (closeIdx !== -1) {
        const strText = codePart.substring(0, closeIdx + 1);
        state.inTemplateStr = false;
        return `<span class="token-string">${escapeHtml(strText)}</span>` + highlightCodeLine(codePart.substring(closeIdx + 1), state);
      } else {
        return `<span class="token-string">${escapeHtml(codePart)}</span>`;
      }
    }

    let lastIndex = 0;
    let html = '';
    JS_TOKEN_REGEX.lastIndex = 0;
    let match;

    while ((match = JS_TOKEN_REGEX.exec(codePart)) !== null) {
      if (match.index > lastIndex) {
        html += escapeHtml(codePart.substring(lastIndex, match.index));
      }

      const [full, comment, str, kw, bool, builtin, num, func, prop, bracket, op] = match;

      if (comment) {
        if (comment.startsWith('/*') && !comment.endsWith('*/')) {
          state.inBlockComment = true;
        }
        html += `<span class="token-comment">${escapeHtml(comment)}</span>`;
      } else if (str) {
        if (str.startsWith('`') && (str.length === 1 || !str.endsWith('`') || str.endsWith('\\`'))) {
          state.inTemplateStr = true;
        }
        html += `<span class="token-string">${escapeHtml(str)}</span>`;
      } else if (kw) {
        html += `<span class="token-keyword">${escapeHtml(kw)}</span>`;
      } else if (bool) {
        html += `<span class="token-boolean">${escapeHtml(bool)}</span>`;
      } else if (builtin) {
        html += `<span class="token-builtin">${escapeHtml(builtin)}</span>`;
      } else if (num) {
        html += `<span class="token-number">${escapeHtml(num)}</span>`;
      } else if (func) {
        html += `<span class="token-func">${escapeHtml(func)}</span>`;
      } else if (prop) {
        html += `<span class="token-prop">${escapeHtml(prop)}</span>`;
      } else if (bracket) {
        let bClass = 'token-bracket-1';
        if ('{([<'.includes(bracket)) {
          bClass = `token-bracket-${(state.bracketDepth % 3) + 1}`;
          state.bracketDepth++;
        } else {
          state.bracketDepth = Math.max(0, state.bracketDepth - 1);
          bClass = `token-bracket-${(state.bracketDepth % 3) + 1}`;
        }
        html += `<span class="${bClass}">${escapeHtml(bracket)}</span>`;
      } else if (op) {
        html += `<span class="token-operator">${escapeHtml(op)}</span>`;
      } else {
        html += escapeHtml(full);
      }

      lastIndex = JS_TOKEN_REGEX.lastIndex;
    }

    if (lastIndex < codePart.length) {
      html += escapeHtml(codePart.substring(lastIndex));
    }

    return html;
  }

  function highlightJS(code) {
    if (!code) return '';
    const lines = code.split('\n');
    const state = { inBlockComment: false, inTemplateStr: false, bracketDepth: 0 };

    const highlightedLines = lines.map(line => {
      const indentMatch = line.match(/^([ \t]+)/);
      const indentStr = indentMatch ? indentMatch[1] : '';
      const codePart = line.substring(indentStr.length);
      const indentHtml = renderIndentGuides(indentStr);
      const codeHtml = highlightCodeLine(codePart, state);
      return indentHtml + codeHtml;
    });

    return highlightedLines.join('\n');
  }

  // Tự động căn chỉnh và thụt lề chuẩn (Auto Format / Beautify JS & ECharts Options)
  function formatJSCode(code) {
    if (!code || typeof code !== 'string') return code;

    const lines = code.split('\n');
    let indentLevel = 0;
    const indentStr = '  '; // 2 spaces chuẩn ECharts
    const resultLines = [];
    let inMultiComment = false;

    for (let i = 0; i < lines.length; i++) {
      let line = lines[i].trim();

      // Giữ dòng trống hợp lý (tối đa 1 dòng trống liên tiếp)
      if (!line) {
        if (resultLines.length > 0 && resultLines[resultLines.length - 1] !== '') {
          resultLines.push('');
        }
        continue;
      }

      // Xử lý comment nhiều dòng /* ... */
      if (line.startsWith('/*')) inMultiComment = true;
      if (inMultiComment) {
        resultLines.push(indentStr.repeat(indentLevel) + line);
        if (line.includes('*/')) inMultiComment = false;
        continue;
      }

      // Đếm số ngoặc đóng ở đầu dòng để thụt lùi ngay dòng này
      let leadingClosers = 0;
      for (let c of line) {
        if (c === '}' || c === ']' || c === ')') {
          leadingClosers++;
        } else if (c === ' ' || c === '\t') {
          continue;
        } else {
          break;
        }
      }

      let currentIndent = Math.max(0, indentLevel - leadingClosers);

      // Đếm thay đổi indent cho các dòng sau (bỏ qua string và single comment)
      let delta = 0;
      let inStr = null;
      let isEscaped = false;

      for (let j = 0; j < line.length; j++) {
        const ch = line[j];

        if (inStr) {
          if (ch === inStr && !isEscaped) inStr = null;
          isEscaped = (ch === '\\' && !isEscaped);
          continue;
        }

        if (ch === '"' || ch === "'" || ch === '`') {
          inStr = ch;
          isEscaped = false;
          continue;
        }

        if (ch === '/' && line[j + 1] === '/') {
          break; // Comment 1 dòng, bỏ qua phần còn lại
        }

        if (ch === '{' || ch === '[' || ch === '(') {
          delta++;
        } else if (ch === '}' || ch === ']' || ch === ')') {
          delta--;
        }
      }

      resultLines.push(indentStr.repeat(currentIndent) + line);
      indentLevel = Math.max(0, indentLevel + delta);
    }

    return resultLines.join('\n');
  }

  function updateEditor() {
    if (!codeEditor) return;
    const code = codeEditor.value;

    // 1. Cập nhật Syntax Highlighting thời gian thực
    if (codeHighlightContent) {
      codeHighlightContent.innerHTML = highlightJS(code) + (code.endsWith('\n') ? ' ' : '');
    }

    // 2. Cập nhật Số dòng
    if (lineNumbers) {
      const lines = code.split('\n').length;
      let numbersHtml = '';
      for (let i = 1; i <= lines; i++) {
        numbersHtml += i + '<br>';
      }
      lineNumbers.innerHTML = numbersHtml;
    }
  }

  // Tương thích ngược
  function updateLineNumbers() {
    updateEditor();
  }

  // Đồng bộ cuộn giữa Line Numbers, Syntax Highlight và Textarea
  if (codeEditor) {
    codeEditor.addEventListener('scroll', () => {
      if (lineNumbers) {
        lineNumbers.scrollTop = codeEditor.scrollTop;
      }
      if (codeHighlight) {
        codeHighlight.scrollTop = codeEditor.scrollTop;
        codeHighlight.scrollLeft = codeEditor.scrollLeft;
      }
    });

    codeEditor.addEventListener('input', updateEditor);

    // Xử lý phím bấm thông minh: Tab thụt lề, Enter auto-indent, đóng ngoặc thông minh
    codeEditor.addEventListener('keydown', (e) => {
      // 1. Phím Tab: thụt lề 2 spaces
      if (e.key === 'Tab') {
        e.preventDefault();
        const start = codeEditor.selectionStart;
        const end = codeEditor.selectionEnd;
        if (e.shiftKey) {
          // Bỏ thụt lề (Shift + Tab)
          const val = codeEditor.value;
          const lineStart = val.lastIndexOf('\n', start - 1) + 1;
          if (val.substring(lineStart, lineStart + 2) === '  ') {
            codeEditor.value = val.substring(0, lineStart) + val.substring(lineStart + 2);
            codeEditor.selectionStart = codeEditor.selectionEnd = Math.max(lineStart, start - 2);
            updateEditor();
          }
        } else {
          codeEditor.value = codeEditor.value.substring(0, start) + '  ' + codeEditor.value.substring(end);
          codeEditor.selectionStart = codeEditor.selectionEnd = start + 2;
          updateEditor();
        }
      } 
      // 2. Phím Enter: Tự động giữ và tăng thụt lề thông minh theo ngữ cảnh
      else if (e.key === 'Enter') {
        if (e.ctrlKey || e.metaKey) {
          e.preventDefault();
          compileAndRenderChart();
          return;
        }

        const start = codeEditor.selectionStart;
        const val = codeEditor.value;
        const lineStart = val.lastIndexOf('\n', start - 1) + 1;
        const currentLine = val.substring(lineStart, start);
        const matchIndent = currentLine.match(/^(\s*)/);
        let indent = matchIndent ? matchIndent[1] : '';

        // Nếu dòng trước kết thúc bằng '{' hoặc '[' thì tự động tăng thêm 2 spaces
        const trimmed = currentLine.trim();
        const shouldAddIndent = trimmed.endsWith('{') || trimmed.endsWith('[');
        if (shouldAddIndent) {
          indent += '  ';
        }

        e.preventDefault();
        const nextChar = val[start];
        const isEnclosed = (trimmed.endsWith('{') && nextChar === '}') || (trimmed.endsWith('[') && nextChar === ']');

        if (isEnclosed) {
          // Tách khối giữa hai ngoặc và tự lùi dòng ngoặc đóng
          const prevIndent = matchIndent ? matchIndent[1] : '';
          const insertText = '\n' + indent + '\n' + prevIndent;
          codeEditor.value = val.substring(0, start) + insertText + val.substring(start);
          codeEditor.selectionStart = codeEditor.selectionEnd = start + indent.length + 1;
        } else {
          codeEditor.value = val.substring(0, start) + '\n' + indent + val.substring(start);
          codeEditor.selectionStart = codeEditor.selectionEnd = start + indent.length + 1;
        }
        updateEditor();
      }
      // 3. Tự động đóng cặp ngoặc khi gõ {, [, (, ", ', `
      else if (e.key === '{' || e.key === '[' || e.key === '(' || e.key === '"' || e.key === "'" || e.key === '`') {
        const pairs = { '{': '}', '[': ']', '(': ')', '"': '"', "'": "'", '`': '`' };
        const closer = pairs[e.key];
        const start = codeEditor.selectionStart;
        const end = codeEditor.selectionEnd;
        if (closer && start === end) {
          e.preventDefault();
          codeEditor.value = codeEditor.value.substring(0, start) + e.key + closer + codeEditor.value.substring(end);
          codeEditor.selectionStart = codeEditor.selectionEnd = start + 1;
          updateEditor();
        }
      }
      // 4. Phím tắt Căn chỉnh code: Shift + Alt + F hoặc Ctrl + Alt + L
      else if ((e.shiftKey && e.altKey && (e.key === 'F' || e.key === 'f')) || 
               (e.ctrlKey && e.altKey && (e.key === 'L' || e.key === 'l'))) {
        e.preventDefault();
        if (btnFormatCode) btnFormatCode.click();
      }
    });
  }

  // =========================================================================
  // 4. TRÌNH BIÊN DỊCH LIVE JAVASCRIPT & RENDER BIỂU ĐỒ
  // =========================================================================
  function printLog(msg, type = 'info') {
    if (!consoleBody) return;
    const time = new Date().toLocaleTimeString('vi-VN');
    const div = document.createElement('div');
    div.className = `viz-log-${type}`;
    div.innerHTML = `<span style="opacity: 0.5;">[${time}]</span> ${msg}`;
    consoleBody.appendChild(div);
    consoleBody.scrollTop = consoleBody.scrollHeight;
  }

  function compileAndRenderChart() {
    if (!myChart || !codeEditor) return;

    const code = codeEditor.value;
    const startTime = performance.now();

    try {
      // Mock Console chuyên biệt
      const customConsole = {
        log: (...args) => {
          printLog(args.map(a => typeof a === 'object' ? JSON.stringify(a) : String(a)).join(' '), 'info');
          console.log(...args);
        },
        warn: (...args) => {
          printLog('⚠️ ' + args.map(a => String(a)).join(' '), 'warn');
          console.warn(...args);
        },
        error: (...args) => {
          printLog('❌ ' + args.map(a => String(a)).join(' '), 'error');
          console.error(...args);
        }
      };

      // Dữ liệu dùng cho biến 'data': Ưu tiên dữ liệu từ file vừa đẩy lên, nếu chưa nạp file thì có dữ liệu mẫu
      const currentData = (rawRows && rawRows.length > 0) ? rawRows : DEFAULT_SAMPLE_DATA;
      window.data = currentData;
      window.rawRows = currentData;

      // Thực thi code an toàn qua new Function - Biến 'data' được truyền làm tham số ĐẦU TIÊN
      const runner = new Function(
        'data',
        'rawRows',
        'files',
        'myChart',
        'echarts',
        'XLSX',
        'console',
        code
      );

      const resultOption = runner(currentData, currentData, files, myChart, echarts, window.XLSX, customConsole);

      if (resultOption && typeof resultOption === 'object') {
        myChart.clear();
        // Nền mặc định luôn là màu trắng (#ffffff). Nếu code trong ECharts có đặt backgroundColor thì giữ nguyên màu code chỉ định.
        if (!resultOption.backgroundColor) {
          resultOption.backgroundColor = '#ffffff';
        }
        myChart.setOption(resultOption, true);
        const duration = (performance.now() - startTime).toFixed(1);
        
        if (renderTimeEl) renderTimeEl.textContent = `${duration} ms`;
        
        // Đếm tổng số điểm dữ liệu
        let totalPoints = 0;
        if (resultOption.series && Array.isArray(resultOption.series)) {
          resultOption.series.forEach(s => {
            if (s.data && Array.isArray(s.data)) totalPoints += s.data.length;
          });
        }
        if (dataCountEl) dataCountEl.textContent = `${totalPoints} điểm`;

        printLog(`✅ Biên dịch & vẽ biểu đồ thành công (${duration} ms).`, 'success');
      } else {
        // Trường hợp người dùng gọi trực tiếp myChart.setOption(option) bên trong code
        const duration = (performance.now() - startTime).toFixed(1);
        if (renderTimeEl) renderTimeEl.textContent = `${duration} ms`;
        printLog(`✅ Code đã thực thi trực tiếp trên myChart (${duration} ms).`, 'success');
      }

    } catch (err) {
      printLog(`❌ Lỗi biên dịch JavaScript: ${err.message}`, 'error');
      console.error(err);
    }
  }

  // =========================================================================
  // 5. NẠP VÀ XỬ LÝ FILE EXCEL / CSV (KÍCH HOẠT BIẾN DATA NGAY TỨC THÌ)
  // =========================================================================
  function handleExcelFile(file) {
    if (!file) return;
    currentFileName = file.name;
    const reader = new FileReader();

    reader.onload = function (e) {
      try {
        const dataBuffer = new Uint8Array(e.target.result);
        const workbook = XLSX.read(dataBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        rawRows = jsonRows;
        files[currentFileName] = jsonRows;
        window.data = jsonRows;
        window.rawRows = jsonRows;

        const cols = jsonRows.length > 0 ? Object.keys(jsonRows[0]) : [];

        // Cập nhật thông tin nạp và chỉ dẫn biến 'data'
        if (dataMetaEl) {
          dataMetaEl.innerHTML = `
            <div style="display: flex; flex-direction: column; gap: 8px;">
              <div>
                <b>📁 File đã nạp:</b> <span style="color: #38bdf8; font-weight: 700;">${currentFileName}</span> | 
                <b>Sheet:</b> ${firstSheetName} | 
                <b>Số dòng:</b> <span style="color: #34d399; font-weight: 700;">${jsonRows.length.toLocaleString('vi-VN')} dòng</span> | 
                <b>Số cột:</b> ${cols.length} cột
              </div>
              <div style="background: rgba(56, 189, 248, 0.12); border: 1px solid rgba(56, 189, 248, 0.35); border-radius: 6px; padding: 8px 12px; font-size: 11.5px; color: var(--text-main);">
                💡 <b>SỬ DỤNG TRỰC TIẾP BIẾN <code style="color: #38bdf8; font-size: 13px; font-weight: bold; background: rgba(0,0,0,0.3); padding: 2px 7px; border-radius: 4px;">data</code>:</b><br>
                Toàn bộ dữ liệu của file đã được truyền thẳng vào biến <code>data</code>. Trong code bạn chỉ cần gọi trực tiếp <code>data.map(r => r["${cols[0] || 'Cột'}"])</code>.
              </div>
              <div style="display: flex; align-items: center; gap: 6px; flex-wrap: wrap;">
                <span style="font-size: 11px; color: var(--text-muted);">Bấm để copy tên cột:</span>
                ${cols.map(c => `<button type="button" class="btn-viz" style="padding: 2px 7px; font-size: 10.5px;" title="Bấm để copy tên cột '${c}'" onclick="navigator.clipboard.writeText('${c.replace(/'/g, "\\'")}'); alert('Đã copy tên cột: ${c}');">📋 ${c}</button>`).join('')}
              </div>
            </div>
          `;
        }

        // Kích hoạt nút Tự Động Vẽ Nhanh Từ Biến "data"
        if (btnAutoPlotData) {
          btnAutoPlotData.style.display = 'flex';
          const colX = cols[0] || 'Cột 1';
          const colY = cols[1] || cols[0] || 'Cột 2';

          btnAutoPlotData.onclick = () => {
            codeEditor.value = `// 📊 TỰ ĐỘNG VẼ BIỂU ĐỒ TRỰC TIẾP TỪ BIẾN 'data'
// File: ${currentFileName} (${jsonRows.length.toLocaleString('vi-VN')} dòng)
// Bạn chỉ cần gọi trực tiếp biến 'data', không cần khai báo lại:

const labelCol = "${colX}";
const valueCol = "${colY}";

// Lấy danh sách nhãn trục X và chuyển đổi số liệu trục Y an toàn
const categories = data.map(r => String(r[labelCol] || ''));
const values = data.map(r => {
  const v = r[valueCol];
  return typeof v === 'number' ? v : parseFloat(String(v || '').replace(/[^0-9.-]/g, '')) || 0;
});

const option = {
  title: {
    text: \`Thống Kê: \${valueCol} Theo \${labelCol}\`,
    subtext: 'Dữ liệu trực tiếp từ biến data (${currentFileName})',
    left: 'center'
  },
  tooltip: {
    trigger: 'axis',
    axisPointer: { type: 'shadow' }
  },
  grid: {
    left: '3%', right: '4%', bottom: '14%', top: '75px',
    containLabel: true
  },
  xAxis: {
    type: 'category',
    data: categories,
    axisLabel: { interval: 0, rotate: categories.length > 8 ? 25 : 0 }
  },
  yAxis: {
    type: 'value',
    name: valueCol,
    splitLine: { lineStyle: { type: 'dashed', opacity: 0.25 } }
  },
  series: [
    {
      name: valueCol,
      type: 'bar',
      barMaxWidth: 35,
      itemStyle: {
        borderRadius: [6, 6, 0, 0],
        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
          { offset: 0, color: '#38bdf8' },
          { offset: 1, color: '#2563eb' }
        ])
      },
      data: values,
      markPoint: { data: [{ type: 'max', name: 'Đỉnh' }, { type: 'min', name: 'Đáy' }] },
      markLine: { data: [{ type: 'average', name: 'Trung bình' }] }
    }
  ]
};

return option;`;

            updateLineNumbers();
            // Tự động chuyển về Tab Code và vẽ biểu đồ ngay lập tức
            if (tabBtns && tabBtns[0]) tabBtns[0].click();
            compileAndRenderChart();
            printLog(`⚡ Đã tự động chèn code và vẽ biểu đồ từ biến 'data' cho file "${currentFileName}"!`, 'success');
          };
        }

        // Render bảng xem trước 5 dòng đầu
        renderPreviewTable(jsonRows.slice(0, 5));

        printLog(`📊 Đã nạp thành công file "${currentFileName}" (${jsonRows.length.toLocaleString('vi-VN')} dòng). Bạn có thể gọi trực tiếp biến "data" trong code!`, 'success');

      } catch (error) {
        printLog(`❌ Lỗi đọc file Excel/CSV: ${error.message}`, 'error');
      }
    };

    reader.readAsArrayBuffer(file);
  }

  function renderPreviewTable(rows) {
    if (!previewTableContainer || rows.length === 0) return;
    const cols = Object.keys(rows[0]);
    let html = `
      <table class="viz-preview-table">
        <thead>
          <tr>
            ${cols.map(c => `<th>${c}</th>`).join('')}
          </tr>
        </thead>
        <tbody>
          ${rows.map(r => `
            <tr>
              ${cols.map(c => `<td>${r[c] !== undefined ? r[c] : ''}</td>`).join('')}
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;
    previewTableContainer.innerHTML = html;
  }

  // Drag & drop file
  if (dropZone && fileInput) {
    dropZone.onclick = () => fileInput.click();
    fileInput.onchange = (e) => {
      if (e.target.files && e.target.files[0]) {
        handleExcelFile(e.target.files[0]);
      }
    };

    dropZone.ondragover = (e) => {
      e.preventDefault();
      dropZone.classList.add('dragover');
    };
    dropZone.ondragleave = () => dropZone.classList.remove('dragover');
    dropZone.ondrop = (e) => {
      e.preventDefault();
      dropZone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        handleExcelFile(e.dataTransfer.files[0]);
      }
    };
  }

  // =========================================================================
  // 6. QUẢN LÝ MODAL MẪU BIỂU ĐỒ (PRESETS)
  // =========================================================================
  function renderPresets() {
    const grid = document.getElementById('vizPresetGrid');
    if (!grid) return;
    grid.innerHTML = CHART_PRESETS.map(p => `
      <div class="preset-card" data-preset-id="${p.id}">
        <div class="preset-card-header">
          <span class="preset-card-icon">${p.icon}</span>
          <span class="preset-card-title">${p.name}</span>
        </div>
        <div class="preset-tag">${p.category}</div>
        <div class="preset-card-desc">${p.desc}</div>
      </div>
    `).join('');

    grid.querySelectorAll('.preset-card').forEach(card => {
      card.onclick = () => {
        const id = card.getAttribute('data-preset-id');
        const preset = CHART_PRESETS.find(p => p.id === id);
        if (preset && codeEditor) {
          codeEditor.value = preset.code;
          updateLineNumbers();
          compileAndRenderChart();
          if (presetModal) presetModal.style.display = 'none';
          printLog(`Đã tải mẫu biểu đồ: "${preset.name}".`, 'info');
        }
      };
    });
  }

  // =========================================================================
  // 7. TRÌNH TẠO PROMPT AI CHẶT CHẼ CHO NON-CODE / LOW-CODE
  // =========================================================================
  const btnOpenFullPage = document.getElementById('btnOpenFullPage');
  const aiInputDataDesc = document.getElementById('aiInputDataDesc');
  const aiBusinessLogic = document.getElementById('aiBusinessLogic');
  const aiExpectedOutput = document.getElementById('aiExpectedOutput');
  const aiPromptFileBadge = document.getElementById('aiPromptFileBadge');
  const aiPromptColsBadge = document.getElementById('aiPromptColsBadge');

  function updateAiPromptModalUI() {
    if (aiPromptFileBadge) {
      if (currentFileName && rawRows && rawRows.length > 0) {
        aiPromptFileBadge.textContent = `📁 ${currentFileName} (${rawRows.length.toLocaleString('vi-VN')} dòng)`;
        aiPromptFileBadge.style.color = '#34d399';
      } else {
        aiPromptFileBadge.textContent = '⚡ Dùng data mẫu mặc định';
        aiPromptFileBadge.style.color = '#f59e0b';
      }
    }

    if (aiPromptColsBadge) {
      const sampleRows = (rawRows && rawRows.length > 0) ? rawRows : DEFAULT_SAMPLE_DATA;
      const cols = sampleRows[0] ? Object.keys(sampleRows[0]) : [];
      aiPromptColsBadge.innerHTML = `
        <span style="color: #60a5fa; font-weight: 600;">Cột khả dụng (${cols.length}):</span> 
        ${cols.map(c => `<code style="background: rgba(255,255,255,0.08); color: #e2e8f0; padding: 1px 5px; border-radius: 3px; font-size: 10px; margin: 1px;">${c}</code>`).join(' ')}
      `;
    }
  }

  function generateEchartsAiPrompt() {
    const activeData = (rawRows && rawRows.length > 0) ? rawRows : DEFAULT_SAMPLE_DATA;
    const isRealFile = Boolean(currentFileName && rawRows && rawRows.length > 0);
    const cols = activeData[0] ? Object.keys(activeData[0]) : [];
    const sampleRows = activeData.slice(0, 3);

    const userGhiChu = (aiInputDataDesc ? aiInputDataDesc.value.trim() : "") || "[Không có ghi chú thêm]";
    const userLogic = (aiBusinessLogic ? aiBusinessLogic.value.trim() : "") || "[Tự động gom nhóm theo danh mục phù hợp và tính tổng các chỉ số số liệu]";
    const userOutput = (aiExpectedOutput ? aiExpectedOutput.value.trim() : "") || "[Biểu đồ cột dải màu gradient sang trọng, tooltip chi tiết, đơn vị VNĐ]";

    const systemPrompt = `Bạn là Kỹ Sư Lập Trình Dữ Liệu JavaScript Cấp Cao (Senior Data Specialist) & Chuyên Gia Apache ECharts.
Hãy viết một đoạn mã Vanilla JavaScript hoàn chỉnh chạy trực tiếp trong trình duyệt để xử lý dữ liệu và cấu hình biểu đồ Apache ECharts theo đúng yêu cầu dưới đây.

---
## 1. MÔI TRƯỜNG THỰC THI (BROWSER RUNTIME)
Code sẽ chạy trực tiếp qua \`new Function('data', 'rawRows', 'files', 'myChart', 'echarts', 'XLSX', 'console', <code>)\`.
- Biến toàn cục \`data\`: ĐÃ CÓ SẴN trong scope, chứa toàn bộ ${activeData.length.toLocaleString('vi-VN')} dòng dữ liệu từ file Excel (Mảng các JSON Object).
- Biến \`myChart\`: Instance ECharts đang hoạt động.
- Biến toàn cục \`echarts\`: Thư viện Apache ECharts v5.5.
- Biến \`console\`: In trực tiếp lên Terminal của công cụ.

---
## 2. DỮ LIỆU ĐẦU VÀO (INPUT DATA CONTEXT)
- **Nguồn dữ liệu**: ${isRealFile ? `File Excel thực tế "${currentFileName}"` : 'Dữ liệu mẫu ban đầu'} (${activeData.length.toLocaleString('vi-VN')} dòng).
- **YÊU CẦU BẮT BUỘC VỚI BIẾN DATA**:
  Code của bạn PHẢI trỏ trực tiếp vào biến \`data\` (ví dụ: \`data.forEach(r => ...)\` hoặc \`data.map(r => ...)\`). TUYỆT ĐỐI KHÔNG định nghĩa lại biến \`data\`, KHÔNG tự bịa mảng dữ liệu giả định.
- **Danh sách cột chính xác 100% trong file (${cols.length} cột)**:
${JSON.stringify(cols, null, 2)}

- **Mẫu 3 dòng dữ liệu đầu tiên (Quan sát kiểu dữ liệu của từng trường)**:
\`\`\`json
${JSON.stringify(sampleRows, null, 2)}
\`\`\`
- **Ghi chú thêm về dữ liệu**: ${userGhiChu}

---
## 3. YÊU CẦU XỬ LÝ NGHIỆP VỤ (BUSINESS LOGIC)
${userLogic}

---
## 4. KẾT QUẢ ĐẦU RA MONG MUỐN (EXPECTED OUTPUT & CHART TYPE)
${userOutput}

---
## 5. BỘ QUY TẮC BẮT BUỘC ĐỂ CODE CHẠY 100% KHÔNG LỖI (CRITICAL RULES)
1. **TRUY CẬP CỘT BẰNG NGOẶC VUÔNG (SAFE ACCESS)**:
   - BẮT BUỘC dùng cú pháp ngoặc vuông: \`row["Tên Cột"]\` (ví dụ: \`row["${cols[0] || 'Tên Cột'}"]\`).
   - TUYỆT ĐỐI KHÔNG dùng cú pháp dấu chấm \`row.Tên Cột\` vì tên cột có thể có khoảng trắng hoặc dấu tiếng Việt.
   - Luôn kiểm tra null/undefined trước khi gọi hàm chuỗi: \`String(row["Tên Cột"] || "").trim()\`.

2. **ÉP KIỂU SỐ AN TOÀN (CLEAN & PARSE NUMBER)**:
   - Dữ liệu Excel có thể là chuỗi chứa dấu phẩy phân cách nghìn hoặc null/khoảng trắng.
   - BẮT BUỘC tạo và sử dụng hàm parse số an toàn:
     \`const parseNum = v => typeof v === 'number' ? (isNaN(v) ? 0 : v) : parseFloat(String(v || '').replace(/[^0-9.-]/g, '')) || 0;\`

3. **NỀN BIỂU ĐỒ & THẨM MỸ CAO CẤP**:
   - NỀN BIỂU ĐỒ BẮT BUỘC LÀ MÀU TRẮNG: Cấu hình \`backgroundColor: '#ffffff'\`.
   - Tooltip phong phú: Sử dụng \`trigger: 'axis'\` (hoặc \`trigger: 'item'\`), định dạng số đẹp bằng \`toLocaleString('vi-VN')\` kèm đơn vị tiền VNĐ / số lượng / tỷ lệ %.
   - Bo góc cột nhẹ nhàng \`borderRadius: [6, 6, 0, 0]\`, màu sắc Gradient hiện đại (xanh MB Bank, Cyan, Emerald...).
   - Bổ sung \`grid: { left: '3%', right: '4%', bottom: '12%', top: '75px', containLabel: true }\`.
   - Bổ sung \`toolbox: { feature: { saveAsImage: { title: 'Tải ảnh PNG' }, dataView: { readOnly: true }, restore: {} } }\`.
   - Nếu danh mục trục X dài hoặc nhiều hơn 6 mục, tự động xoay nhãn \`axisLabel: { interval: 0, rotate: 20 }\` hoặc thêm thanh trượt \`dataZoom\`.

4. **BẮT BUỘC COMMENT ĐẦY ĐỦ ĐỂ DỄ BẢO TRÌ (CODE MAINTENANCE & CLEAN CODE)**:
   - Mã nguồn BẮT BUỘC phải có comment tiếng Việt chi tiết, rõ ràng từng phân đoạn chức năng:
     * // ==========================================
     * // 1. TRÍCH XUẤT VÀ LÀM SẠCH DỮ LIỆU ĐẦU VÀO: Đọc mảng 'data', kiểm tra rỗng, ép kiểu Number an toàn.
     * // 2. GOM NHÓM (AGGREGATE) & TÍNH TOÁN NGHIỆP VỤ: Diễn giải công thức tính toán (SUM, AVG, COUNT, tỷ lệ %), logic Map/Reduce, sắp xếp.
     * // 3. THIẾT LẬP CẤU HÌNH OPTION ECHARTS: Chú thích màu sắc gradient, tooltip format tiền tệ VNĐ/%, trục X/Y, bo góc cột, toolbox.
     * // 4. HIỆU ỨNG TƯƠNG TÁC CLICK & DRILL-DOWN (NẾU CÓ): Chú thích rõ cách bắt sự kiện click, chống double event, xử lý logic tương tác.
   - Đặt tên biến tường minh, phản ánh đúng ý nghĩa nghiệp vụ (ví dụ: totalRevenue, salesByCategory, topBranches, categories), TUYỆT ĐỐI KHÔNG dùng tên tắt khó hiểu (a, b, t1, x).

5. **QUY CHUẨN HIỆU ỨNG CLICK TƯƠNG TÁC TRÊN BIỂU ĐỒ (CHART CLICK EVENTS)**:
   Khi biểu đồ có yêu cầu tương tác click hoặc khoan sâu (drill-down), BẮT BUỘC tuân thủ các quy tắc sau:
   - Click vào phần tử dữ liệu (Cột, Lát bánh Pie, Điểm đường Line, Thanh ngang...):
     * Dùng cú pháp:
       \`\`\`javascript
       // Chống nhân bản sự kiện (double firing) khi biểu đồ render lại hoặc khi lọc dữ liệu
       if (myChart && myChart.off) {
         myChart.off('click');
         myChart.on('click', function(params) {
           // params.name: nhãn danh mục (vd: "Hà Nội", "Miền Bắc", "Tháng 05")
           // params.value: giá trị số liệu của phần tử
           // params.seriesName: tên chuỗi dữ liệu
           // params.dataIndex: chỉ số vị trí dữ liệu
           console.log('Đã click vào phần tử:', params.name, params.value);
           
           // Xử lý logic tương tác (ví dụ: drill-down, hiển thị modal chi tiết, hoặc lọc dữ liệu)
         });
       }
       \`\`\`
   - Hiệu ứng thị giác khi di chuột & click (Visual Feedback):
     * Đặt con trỏ chuột dạng bàn tay: Thêm \`cursor: 'pointer'\` trong cấu hình series hoặc \`myChart.getZr().setCursorStyle('pointer')\`.
     * Highlight phần tử đang chọn: Có thể dùng \`selectedMode: 'single'\` trên series, hoặc gọi \`myChart.dispatchAction({ type: 'highlight', seriesIndex: params.seriesIndex, dataIndex: params.dataIndex })\`.
   - Nút chức năng (Graphic Buttons):
     * Dùng mảng \`graphic: [...] \` của ECharts với \`type: 'group'\`, bên trong có \`rect\` (nền nút) và \`text\` (chữ), gắn \`cursor: 'pointer'\` và sự kiện \`onclick: function() { ... }\` để chuyển đổi kiểu biểu đồ (Bar/Line) hoặc reset zoom.
     * Hoặc dùng HTML DOM Overlay qua \`const dom = myChart.getDom();\` chèn thẻ button HTML thật có class CSS nổi góc trên.

6. **ĐỊNH DẠNG ĐẦU RA BẮT BUỘC**:
   - BẮT BUỘC kết thúc bằng câu lệnh: \`return option;\` (hoặc gọi \`myChart.setOption(option);\`).
   - CHỈ TRẢ VỀ DUY NHẤT một khối mã JavaScript trong \`\`\`javascript ... \`\`\`.
   - KHÔNG viết lời chào hỏi, cảm ơn hay giải thích dài dòng bên ngoài khối mã.
   - KHÔNG dùng cú pháp \`import\`, \`export\`, \`require\` (vì code chạy trực tiếp trong browser).`;

    if (aiPromptOutput) {
      aiPromptOutput.value = systemPrompt.trim();
    }
  }

  // =========================================================================
  // =========================================================================
  // 8. TÍNH NĂNG MỞ FULL TRANG WEB MỚI (FULL PAGE VIEW - HỖ TRỢ ĐẦY ĐỦ TƯƠNG TÁC)
  // =========================================================================
  function openChartInNewTab() {
    if (!myChart) return;
    const currentOpt = (myChart.getOption && myChart.getOption()) || {};
    const chartTitle = (currentOpt.title && currentOpt.title[0] && currentOpt.title[0].text) || 'ECharts Interactive Visualizer';
    const bg = currentOpt.backgroundColor || '#ffffff';

    // Lấy mã code hiện tại và bộ dữ liệu đang dùng
    const userCode = (codeEditor && codeEditor.value) ? codeEditor.value : '';
    const currentData = (rawRows && rawRows.length > 0) ? rawRows : DEFAULT_SAMPLE_DATA;
    const currentFiles = files || {};

    // Chuẩn bị dữ liệu an toàn để nhúng vào thẻ script (tránh đóng thẻ sớm)
    const safeUserCode = JSON.stringify(userCode).replace(/<\/script>/gi, '<\\/script>');
    const safeData = JSON.stringify(currentData).replace(/<\/script>/gi, '<\\/script>');
    const safeFiles = JSON.stringify(currentFiles).replace(/<\/script>/gi, '<\\/script>');
    const safeFallbackOpt = JSON.stringify(currentOpt).replace(/<\/script>/gi, '<\\/script>');

    // Tạo trang HTML độc lập hoàn chỉnh với Live Runtime đầy đủ
    const fullHtml = `<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${chartTitle}</title>
  <!-- Nhúng Apache ECharts & SheetJS XLSX -->
  <script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"><\/script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"><\/script>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    html, body {
      width: 100vw;
      height: 100vh;
      overflow: hidden;
      background: ${bg};
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    #fullChartCanvas {
      width: 100vw;
      height: 100vh;
    }
    .floating-toolbar {
      position: fixed;
      top: 16px;
      right: 20px;
      display: flex;
      align-items: center;
      gap: 8px;
      z-index: 99999;
      background: rgba(15, 23, 42, 0.88);
      backdrop-filter: blur(14px) saturate(160%);
      -webkit-backdrop-filter: blur(14px) saturate(160%);
      padding: 6px 10px;
      border-radius: 12px;
      border: 1px solid rgba(255, 255, 255, 0.16);
      box-shadow: 0 10px 30px -5px rgba(0, 0, 0, 0.45), 0 0 0 1px rgba(255, 255, 255, 0.06);
      transition: box-shadow 0.2s ease, transform 0.15s ease, background 0.2s ease;
      user-select: none;
      touch-action: none;
    }
    .floating-toolbar.is-dragging {
      opacity: 0.95;
      transform: scale(1.02);
      box-shadow: 0 16px 40px rgba(0, 0, 0, 0.55), 0 0 0 2px rgba(56, 189, 248, 0.5);
      cursor: grabbing !important;
    }
    .toolbar-drag-handle {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      color: #94a3b8;
      font-size: 15px;
      cursor: grab;
      padding: 4px 6px;
      border-radius: 6px;
      transition: all 0.15s ease;
      letter-spacing: -2px;
      line-height: 1;
    }
    .toolbar-drag-handle:hover {
      color: #38bdf8;
      background: rgba(255, 255, 255, 0.1);
    }
    .floating-toolbar.is-dragging .toolbar-drag-handle {
      cursor: grabbing;
    }
    .floating-toolbar-items {
      display: flex;
      align-items: center;
      gap: 8px;
      transition: opacity 0.2s ease, max-width 0.25s ease;
      white-space: nowrap;
    }
    .floating-toolbar.collapsed {
      padding: 5px 8px;
      background: rgba(15, 23, 42, 0.8);
      border-radius: 10px;
    }
    .floating-toolbar.collapsed .floating-toolbar-items {
      display: none;
    }
    .floating-btn {
      background: linear-gradient(135deg, #2563eb, #1d4ed8);
      color: #ffffff;
      border: 1px solid rgba(255, 255, 255, 0.15);
      padding: 6px 13px;
      border-radius: 6px;
      font-size: 12px;
      font-weight: 600;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 5px;
      box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
      transition: all 0.18s cubic-bezier(0.16, 1, 0.3, 1);
    }
    .floating-btn:hover {
      background: linear-gradient(135deg, #3b82f6, #2563eb);
      transform: translateY(-1px);
      box-shadow: 0 4px 10px rgba(37, 99, 235, 0.35);
    }
    .floating-btn:active {
      transform: translateY(0);
    }
    .floating-btn-close {
      background: rgba(71, 85, 105, 0.65);
      box-shadow: none;
      border: 1px solid rgba(255, 255, 255, 0.1);
    }
    .floating-btn-close:hover {
      background: #dc2626;
      border-color: #ef4444;
      box-shadow: 0 4px 10px rgba(220, 38, 38, 0.35);
    }
    .floating-btn-toggle {
      background: rgba(255, 255, 255, 0.08);
      color: #e2e8f0;
      border: 1px solid rgba(255, 255, 255, 0.12);
      padding: 5px 10px;
      border-radius: 6px;
      font-size: 12px;
      font-weight: 600;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 5px;
      transition: all 0.18s cubic-bezier(0.16, 1, 0.3, 1);
    }
    .floating-btn-toggle:hover {
      background: rgba(255, 255, 255, 0.18);
      color: #ffffff;
      border-color: rgba(255, 255, 255, 0.25);
      transform: translateY(-1px);
    }
    .floating-btn-toggle:active {
      transform: translateY(0);
    }
    .status-badge {
      font-size: 11px;
      font-weight: 500;
      color: #94a3b8;
      background: rgba(255, 255, 255, 0.05);
      padding: 4px 8px;
      border-radius: 4px;
      border: 1px solid rgba(255, 255, 255, 0.08);
      margin-right: 2px;
    }
  </style>
</head>
<body>
  <div class="floating-toolbar" id="floatingToolbar">
    <!-- Tay cầm kéo thả di chuyển -->
    <div class="toolbar-drag-handle" id="toolbarDragHandle" title="Bấm giữ để di chuyển thanh công cụ quanh màn hình">
      <span>⠿</span>
    </div>
    <div class="floating-toolbar-items" id="toolbarItems">
      <span class="status-badge">⚡ Full Interactive Mode</span>
      <button class="floating-btn" onclick="downloadImage()">📷 Tải Ảnh</button>
      <button class="floating-btn floating-btn-close" onclick="window.close()">✕ Đóng</button>
    </div>
    <button class="floating-btn-toggle" id="btnToggleToolbar" onclick="toggleToolbar()" title="Thu gọn thanh công cụ">
      <span id="toggleIcon">➖</span> <span id="toggleLabel">Thu gọn</span>
    </button>
  </div>
  <div id="fullChartCanvas"></div>
  <script>
    (function() {
      // 1. Dữ liệu nạp từ visualizer
      const data = ${safeData};
      const rawRows = data;
      const files = ${safeFiles};
      const userCode = ${safeUserCode};
      const fallbackOpt = ${safeFallbackOpt};

      // Đưa vào window để tiện debug
      window.data = data;
      window.rawRows = rawRows;
      window.files = files;

      // 2. Khởi tạo Apache ECharts instance
      const chartDom = document.getElementById('fullChartCanvas');
      const myChart = echarts.init(chartDom, null, { renderer: 'canvas' });
      window.myChart = myChart;

      // 3. Thực thi trực tiếp mã JavaScript để giữ trọn vẹn sự kiện onclick, filter và formatter
      let rendered = false;
      if (userCode && userCode.trim()) {
        try {
          const runner = new Function('data', 'rawRows', 'files', 'myChart', 'echarts', 'XLSX', 'console', userCode);
          const resultOption = runner(data, rawRows, files, myChart, window.echarts, window.XLSX, console);
          if (resultOption && typeof resultOption === 'object') {
            if (!resultOption.backgroundColor) {
              resultOption.backgroundColor = ${JSON.stringify(bg)};
            }
            myChart.setOption(resultOption, true);
            rendered = true;
          } else {
            // Nếu người dùng gọi trực tiếp myChart.setOption trong code
            rendered = true;
          }
        } catch (err) {
          console.warn('Không thể chạy lại runtime code, dùng snapshot fallback:', err);
        }
      }

      // 4. Fallback nếu không có code hoặc lỗi
      if (!rendered && fallbackOpt && Object.keys(fallbackOpt).length > 0) {
        myChart.setOption(fallbackOpt);
      }

      // 5. Tự động co giãn theo kích thước cửa sổ
      window.addEventListener('resize', function() {
        if (myChart) myChart.resize();
      });

      // 6. Hàm đóng / mở thanh công cụ
      window.toggleToolbar = function() {
        const tb = document.getElementById('floatingToolbar');
        const icon = document.getElementById('toggleIcon');
        const label = document.getElementById('toggleLabel');
        const btn = document.getElementById('btnToggleToolbar');
        if (!tb) return;
        
        const isCollapsed = tb.classList.toggle('collapsed');
        if (isCollapsed) {
          icon.textContent = '🛠️';
          label.textContent = 'Công cụ';
          btn.title = 'Mở rộng thanh công cụ';
        } else {
          icon.textContent = '➖';
          label.textContent = 'Thu gọn';
          btn.title = 'Thu gọn thanh công cụ';
        }
      };

      // 7. Cơ chế Kéo Thả Di Chuyển Thanh Công Cụ Quanh Màn Hình (Draggable Toolbar)
      const tb = document.getElementById('floatingToolbar');
      let isDragging = false;
      let startX = 0, startY = 0, initialLeft = 0, initialTop = 0;

      function onPointerDown(e) {
        // Không kích hoạt kéo thả nếu người dùng click trực tiếp vào nút bấm
        if (e.target.closest('button')) return;

        isDragging = true;
        const clientX = e.clientX !== undefined ? e.clientX : (e.touches && e.touches[0] ? e.touches[0].clientX : 0);
        const clientY = e.clientY !== undefined ? e.clientY : (e.touches && e.touches[0] ? e.touches[0].clientY : 0);

        startX = clientX;
        startY = clientY;

        const rect = tb.getBoundingClientRect();
        initialLeft = rect.left;
        initialTop = rect.top;

        tb.classList.add('is-dragging');

        // Bỏ thuộc tính right / bottom để điều khiển left / top trực tiếp
        tb.style.right = 'auto';
        tb.style.bottom = 'auto';
        tb.style.left = initialLeft + 'px';
        tb.style.top = initialTop + 'px';

        document.addEventListener('mousemove', onPointerMove);
        document.addEventListener('mouseup', onPointerUp);
        document.addEventListener('touchmove', onPointerMove, { passive: false });
        document.addEventListener('touchend', onPointerUp);
      }

      function onPointerMove(e) {
        if (!isDragging) return;
        if (e.cancelable) e.preventDefault();

        const clientX = e.clientX !== undefined ? e.clientX : (e.touches && e.touches[0] ? e.touches[0].clientX : 0);
        const clientY = e.clientY !== undefined ? e.clientY : (e.touches && e.touches[0] ? e.touches[0].clientY : 0);

        const deltaX = clientX - startX;
        const deltaY = clientY - startY;

        let newLeft = initialLeft + deltaX;
        let newTop = initialTop + deltaY;

        // Giới hạn trong phạm vi màn hình, không cho trôi ra ngoài mép
        const tbWidth = tb.offsetWidth;
        const tbHeight = tb.offsetHeight;
        const padding = 8;

        newLeft = Math.max(padding, Math.min(window.innerWidth - tbWidth - padding, newLeft));
        newTop = Math.max(padding, Math.min(window.innerHeight - tbHeight - padding, newTop));

        tb.style.left = newLeft + 'px';
        tb.style.top = newTop + 'px';
      }

      function onPointerUp() {
        if (!isDragging) return;
        isDragging = false;
        tb.classList.remove('is-dragging');

        document.removeEventListener('mousemove', onPointerMove);
        document.removeEventListener('mouseup', onPointerUp);
        document.removeEventListener('touchmove', onPointerMove);
        document.removeEventListener('touchend', onPointerUp);
      }

      tb.addEventListener('mousedown', onPointerDown);
      tb.addEventListener('touchstart', onPointerDown, { passive: true });

      // Khi resize cửa sổ, giữ toolbar luôn nằm bên trong màn hình
      window.addEventListener('resize', function() {
        if (!tb.style.left) return;
        const rect = tb.getBoundingClientRect();
        const padding = 8;
        let left = rect.left;
        let top = rect.top;

        if (left + rect.width > window.innerWidth - padding) {
          left = Math.max(padding, window.innerWidth - rect.width - padding);
          tb.style.left = left + 'px';
        }
        if (top + rect.height > window.innerHeight - padding) {
          top = Math.max(padding, window.innerHeight - rect.height - padding);
          tb.style.top = top + 'px';
        }
      });

      // 8. Hàm chụp ảnh biểu đồ
      window.downloadImage = function() {
        if (!myChart) return;
        const url = myChart.getDataURL({ type: 'png', pixelRatio: 2, backgroundColor: ${JSON.stringify(bg)} });
        const a = document.createElement('a');
        a.href = url;
        a.download = 'echarts_fullpage_' + Date.now() + '.png';
        a.click();
      };
    })();
  <\/script>
</body>
</html>`;

    const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
    const blobUrl = URL.createObjectURL(blob);
    window.open(blobUrl, '_blank');
    printLog("🌐 Đã mở biểu đồ toàn màn hình với đầy đủ chức năng tương tác nút bấm trong tab mới!", "success");
  }

  // =========================================================================
  // 8. SỰ KIỆN GIAO DIỆN & KHỞI CHẠY
  // =========================================================================
  function setupEvents() {
    // Tab switching
    tabBtns.forEach(btn => {
      btn.onclick = () => {
        tabBtns.forEach(b => b.classList.remove('active'));
        tabContents.forEach(c => c.classList.remove('active'));
        btn.classList.add('active');
        const targetId = btn.getAttribute('data-tab');
        const targetContent = document.getElementById(targetId);
        if (targetContent) targetContent.classList.add('active');
      };
    });

    // Run code
    if (btnRunCode) btnRunCode.onclick = compileAndRenderChart;

    // Reset code to default preset
    if (btnResetCode) {
      btnResetCode.onclick = () => {
        if (confirm("Khôi phục code về mẫu biểu đồ cột ban đầu?")) {
          codeEditor.value = CHART_PRESETS[0].code;
          updateEditor();
          compileAndRenderChart();
        }
      };
    }

    // Tự động căn chỉnh & thụt lề code (Format Code)
    if (btnFormatCode) {
      btnFormatCode.onclick = () => {
        if (!codeEditor) return;
        const originalCode = codeEditor.value;
        const formatted = formatJSCode(originalCode);
        if (formatted !== originalCode) {
          codeEditor.value = formatted;
          updateEditor();
          printLog("✨ Đã tự động căn chỉnh và thụt lề code chuẩn đẹp!", "success");
        } else {
          printLog("✨ Code đã được căn chỉnh chuẩn xác sẵn.", "info");
        }
      };
    }

    // Copy Code
    if (btnCopyCode) {
      btnCopyCode.onclick = () => {
        if (!codeEditor) return;
        navigator.clipboard.writeText(codeEditor.value).then(() => {
          alert("Đã sao chép toàn bộ code JavaScript!");
        });
      };
    }

    // Export PNG
    if (btnExportPng) {
      btnExportPng.onclick = () => {
        if (!myChart) return;
        const currentOpt = (myChart.getOption && myChart.getOption()) || {};
        const bg = currentOpt.backgroundColor || '#ffffff';
        const url = myChart.getDataURL({
          type: 'png',
          pixelRatio: 2,
          backgroundColor: bg
        });
        const a = document.createElement('a');
        a.href = url;
        a.download = `echarts_chart_${Date.now()}.png`;
        a.click();
        printLog("📷 Đã tải ảnh biểu đồ độ phân giải cao về máy!", "success");
      };
    }

    // Clear console
    if (btnClearConsole) {
      btnClearConsole.onclick = () => {
        if (consoleBody) consoleBody.innerHTML = '';
      };
    }

    // Xuất ra Full Trang Web Mới trong Tab Trình Duyệt Độc Lập
    if (btnOpenFullPage) {
      btnOpenFullPage.onclick = openChartInNewTab;
    }

    // Modal Presets
    if (btnOpenPresets && presetModal) {
      btnOpenPresets.onclick = () => {
        renderPresets();
        presetModal.style.display = 'flex';
      };
    }
    if (btnClosePresetModal && presetModal) {
      btnClosePresetModal.onclick = () => presetModal.style.display = 'none';
    }

    // Modal AI Prompt (Chặt chẽ cho non-code / low-code)
    if (btnOpenAiPrompt && aiPromptModal) {
      btnOpenAiPrompt.onclick = () => {
        updateAiPromptModalUI();
        generateEchartsAiPrompt();
        aiPromptModal.style.display = 'flex';
      };
    }
    if (btnCloseAiModal && aiPromptModal) {
      btnCloseAiModal.onclick = () => aiPromptModal.style.display = 'none';
    }
    if (btnGenerateAiPrompt) {
      btnGenerateAiPrompt.onclick = generateEchartsAiPrompt;
    }

    // Tự động cập nhật System Prompt khi người dùng gõ
    [aiInputDataDesc, aiBusinessLogic, aiExpectedOutput].forEach(el => {
      if (el) el.oninput = generateEchartsAiPrompt;
    });

    // Bấm chọn nhanh loại biểu đồ qua chip
    const typeChips = document.querySelectorAll('.ai-type-chip');
    typeChips.forEach(chip => {
      chip.onclick = () => {
        const typeDesc = chip.getAttribute('data-type');
        if (aiExpectedOutput) {
          aiExpectedOutput.value = typeDesc;
          generateEchartsAiPrompt();
        }
      };
    });

    if (btnCopyAiPrompt && aiPromptOutput) {
      btnCopyAiPrompt.onclick = () => {
        navigator.clipboard.writeText(aiPromptOutput.value).then(() => {
          alert("Đã copy System Prompt vào Clipboard! Hãy dán vào ChatGPT / Gemini / Claude để nhận code chuẩn 100%.");
        });
      };
    }

    // Theme Toggle Sync
    if (themeToggleBtn) {
      themeToggleBtn.onclick = () => {
        if (window.ThemeManager && typeof window.ThemeManager.toggleTheme === 'function') {
          window.ThemeManager.toggleTheme();
          setTimeout(() => {
            initChart();
            compileAndRenderChart();
          }, 150);
        }
      };
    }

    // Đóng modal khi bấm ra ngoài
    window.onclick = (e) => {
      if (e.target === presetModal) presetModal.style.display = 'none';
      if (e.target === aiPromptModal) aiPromptModal.style.display = 'none';
    };

    // Kéo giãn thanh Resizer giữa 2 cột
    const resizer = document.getElementById('vizResizer');
    const editorPanel = document.querySelector('.viz-editor-panel');
    if (resizer && editorPanel) {
      let isResizing = false;
      resizer.addEventListener('mousedown', (e) => {
        isResizing = true;
        resizer.classList.add('resizing');
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';
      });

      document.addEventListener('mousemove', (e) => {
        if (!isResizing) return;
        const offsetLeft = e.clientX;
        const totalWidth = window.innerWidth;
        const percentage = (offsetLeft / totalWidth) * 100;
        if (percentage >= 25 && percentage <= 75) {
          editorPanel.style.width = percentage + '%';
          if (myChart) myChart.resize();
        }
      });

      document.addEventListener('mouseup', () => {
        if (isResizing) {
          isResizing = false;
          resizer.classList.remove('resizing');
          document.body.style.cursor = '';
          document.body.style.userSelect = '';
          if (myChart) myChart.resize();
        }
      });
    }
  }

  // Khởi động khi tải xong trang
  window.addEventListener('DOMContentLoaded', () => {
    initChart();
    if (codeEditor) {
      codeEditor.value = CHART_PRESETS[0].code;
      updateEditor();
    }
    setupEvents();
    // Chạy render mẫu mặc định
    setTimeout(compileAndRenderChart, 100);
  });

})();
