import React, { useEffect, useMemo, useRef, useState } from 'react';
import './App.css';
import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';

const USERS_KEY = 'excel_analytics_users';
const SESSION_KEY = 'excel_analytics_session';
const FILES_KEY_PREFIX = 'excel_analytics_files_';
const CHART_COLORS = ['#0f766e', '#f97316', '#2563eb', '#e11d48', '#7c3aed'];

const readJson = (key, fallback) => {
  try {
    const raw = window.localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (error) {
    console.error(`Unable to read ${key}`, error);
    return fallback;
  }
};

const writeJson = (key, value) => {
  window.localStorage.setItem(key, JSON.stringify(value));
};

const loadFiles = (email) => readJson(`${FILES_KEY_PREFIX}${email}`, []);
const saveFiles = (email, files) => writeJson(`${FILES_KEY_PREFIX}${email}`, files);

const formatNumber = (value) => {
  if (typeof value !== 'number' || Number.isNaN(value)) {
    return '0';
  }

  return new Intl.NumberFormat('en-IN', { maximumFractionDigits: 2 }).format(value);
};

const formatDate = (value) => {
  try {
    return new Date(value).toLocaleString();
  } catch {
    return value;
  }
};

const downloadTextFile = (filename, content, type = 'text/plain;charset=utf-8;') => {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
};

const toTableData = (sheetRows) => {
  const [headerRow = [], ...bodyRows] = sheetRows;
  const headers = headerRow.map((header, index) => String(header || `Column ${index + 1}`).trim());

  const rows = bodyRows
    .filter((row) => row.some((cell) => cell !== undefined && cell !== null && `${cell}`.trim() !== ''))
    .map((row, rowIndex) => {
      const record = { __rowId: `${Date.now()}-${rowIndex}` };
      headers.forEach((header, columnIndex) => {
        record[header] = row[columnIndex] ?? '';
      });
      return record;
    });

  return { headers, rows };
};

const getNumericHeaders = (headers, rows) =>
  headers.filter((header) => rows.some((row) => row[header] !== '' && !Number.isNaN(Number(row[header]))));

const buildInsights = (rows, numericHeaders) =>
  numericHeaders.slice(0, 3).map((header) => {
    const values = rows.map((row) => Number(row[header])).filter((value) => !Number.isNaN(value));
    const total = values.reduce((sum, value) => sum + value, 0);
    const average = values.length ? total / values.length : 0;
    const peak = values.length ? Math.max(...values) : 0;

    return {
      label: header,
      total,
      average,
      peak,
    };
  });

const getChartData = (rows, xAxis, yAxis) =>
  rows
    .map((row) => ({
      name: String(row[xAxis] ?? 'Untitled'),
      value: Number(row[yAxis]),
    }))
    .filter((item) => item.name && !Number.isNaN(item.value))
    .slice(0, 12);

const getTopCategory = (rows, categoryHeader, numericHeader) => {
  if (!categoryHeader || !numericHeader) {
    return null;
  }

  const grouped = rows.reduce((accumulator, row) => {
    const key = String(row[categoryHeader] ?? 'Unknown');
    const value = Number(row[numericHeader]);
    if (Number.isNaN(value)) {
      return accumulator;
    }

    accumulator[key] = (accumulator[key] || 0) + value;
    return accumulator;
  }, {});

  const entries = Object.entries(grouped).sort((left, right) => right[1] - left[1]);
  return entries[0] ? { label: entries[0][0], value: entries[0][1] } : null;
};

const buildAiSnapshot = (file) => {
  const primaryMetric = file.numericHeaders[0];
  const secondaryMetric = file.numericHeaders[1];
  const leadCategory = file.headers.find((header) => !file.numericHeaders.includes(header));
  const primaryValues = primaryMetric
    ? file.rows.map((row) => Number(row[primaryMetric])).filter((value) => !Number.isNaN(value))
    : [];
  const total = primaryValues.reduce((sum, value) => sum + value, 0);
  const average = primaryValues.length ? total / primaryValues.length : 0;
  const minimum = primaryValues.length ? Math.min(...primaryValues) : 0;
  const maximum = primaryValues.length ? Math.max(...primaryValues) : 0;
  const change = primaryValues.length > 1 ? primaryValues[primaryValues.length - 1] - primaryValues[0] : 0;
  const trend = change > 0 ? 'upward' : change < 0 ? 'downward' : 'flat';
  const topCategory = getTopCategory(file.rows, leadCategory, primaryMetric);

  return {
    primaryMetric,
    secondaryMetric,
    leadCategory,
    total,
    average,
    minimum,
    maximum,
    change,
    trend,
    topCategory,
  };
};

const generateAiResponse = (file, question) => {
  const prompt = question.trim().toLowerCase();
  const snapshot = buildAiSnapshot(file);

  if (!snapshot.primaryMetric) {
    return 'I can help once the file has at least one numeric column. Right now I recommend uploading data with values like revenue, sales, score, or quantity.';
  }

  if (prompt.includes('trend') || prompt.includes('growing') || prompt.includes('decline')) {
    const first = Number(file.rows[0]?.[snapshot.primaryMetric] ?? 0);
    const last = Number(file.rows[file.rows.length - 1]?.[snapshot.primaryMetric] ?? 0);
    return `The main trend in ${snapshot.primaryMetric} looks ${snapshot.trend}. It moved from ${formatNumber(first)} to ${formatNumber(last)}, with an overall change of ${formatNumber(snapshot.change)}.`;
  }

  if (prompt.includes('top') || prompt.includes('best') || prompt.includes('highest')) {
    if (snapshot.topCategory) {
      return `${snapshot.topCategory.label} is currently the strongest ${snapshot.leadCategory || 'segment'} based on ${snapshot.primaryMetric}, contributing ${formatNumber(snapshot.topCategory.value)}.`;
    }

    return `The highest observed ${snapshot.primaryMetric} value in this file is ${formatNumber(snapshot.maximum)}.`;
  }

  if (prompt.includes('average') || prompt.includes('mean')) {
    return `The average ${snapshot.primaryMetric} is ${formatNumber(snapshot.average)} across ${formatNumber(file.rowCount)} rows.`;
  }

  if (prompt.includes('risk') || prompt.includes('issue') || prompt.includes('problem')) {
    return `The biggest risk is concentration around ${snapshot.primaryMetric}. I would watch rows near the minimum value of ${formatNumber(snapshot.minimum)} and compare them against the stronger rows near ${formatNumber(snapshot.maximum)}.`;
  }

  if (prompt.includes('recommend') || prompt.includes('next') || prompt.includes('chart')) {
    const chartMetric = snapshot.secondaryMetric || snapshot.primaryMetric;
    const chartDimension = snapshot.leadCategory || file.headers[0];
    return `I recommend charting ${chartMetric} by ${chartDimension}. That view should separate strong versus weak segments faster than scanning the raw table.`;
  }

  return `Here is the quick read: ${snapshot.primaryMetric} totals ${formatNumber(snapshot.total)}, averages ${formatNumber(snapshot.average)}, and the overall pattern looks ${snapshot.trend}. ${snapshot.topCategory ? `${snapshot.topCategory.label} is the standout ${snapshot.leadCategory || 'segment'}.` : ''}`.trim();
};

const buildAiCards = (file) => {
  const snapshot = buildAiSnapshot(file);

  if (!snapshot.primaryMetric) {
    return [];
  }

  return [
    {
      title: 'AI Summary',
      body: `${snapshot.primaryMetric} averages ${formatNumber(snapshot.average)} and the overall movement looks ${snapshot.trend}.`,
    },
    {
      title: 'AI Opportunity',
      body: snapshot.topCategory
        ? `${snapshot.topCategory.label} is leading by ${formatNumber(snapshot.topCategory.value)}. This is a good segment to benchmark.`
        : `The highest ${snapshot.primaryMetric} value is ${formatNumber(snapshot.maximum)}.`,
    },
    {
      title: 'AI Recommendation',
      body: `Try plotting ${snapshot.secondaryMetric || snapshot.primaryMetric} against ${snapshot.leadCategory || file.headers[0]} for a cleaner decision view.`,
    },
  ];
};

const buildDemoWorkbook = () => ({
  name: 'demo-sales-dashboard.xlsx',
  sheetName: 'Quarterly Sales',
  rows: [
    ['Month', 'Region', 'Revenue', 'Orders', 'Satisfaction'],
    ['January', 'North', 120000, 420, 4.4],
    ['February', 'North', 132500, 465, 4.6],
    ['March', 'North', 141300, 492, 4.7],
    ['April', 'East', 118400, 401, 4.3],
    ['May', 'East', 126900, 433, 4.5],
    ['June', 'East', 138250, 470, 4.6],
    ['July', 'South', 144100, 508, 4.4],
    ['August', 'South', 152000, 541, 4.8],
    ['September', 'West', 149500, 520, 4.7],
    ['October', 'West', 158600, 566, 4.9],
  ],
});

const IconWrap = ({ children }) => <span className="icon-wrap">{children}</span>;

const UploadIcon = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
    <path d="M12 16V4" />
    <path d="m7 9 5-5 5 5" />
    <path d="M4 16.5v1a2.5 2.5 0 0 0 2.5 2.5h11A2.5 2.5 0 0 0 20 17.5v-1" />
  </svg>
);

const ChartIcon = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
    <path d="M3 3v18h18" />
    <path d="M7 14v3" />
    <path d="M12 9v8" />
    <path d="M17 6v11" />
  </svg>
);

const SparkIcon = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
    <path d="M12 3l1.9 4.6L18.5 9l-4.6 1.4L12 15l-1.9-4.6L5.5 9l4.6-1.4L12 3Z" />
    <path d="M19 15l.9 2.1L22 18l-2.1.9L19 21l-.9-2.1L16 18l2.1-.9L19 15Z" />
  </svg>
);

const LogoutIcon = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
    <path d="M10 17l5-5-5-5" />
    <path d="M15 12H3" />
    <path d="M20 19v-2a2 2 0 0 0-2-2h-4" />
    <path d="M20 5v2a2 2 0 0 1-2 2h-4" />
  </svg>
);

const FileIcon = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
    <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8Z" />
    <path d="M14 3v5h5" />
    <path d="M9 13h6" />
    <path d="M9 17h6" />
  </svg>
);

const Button = ({ children, className = '', variant = 'primary', ...props }) => (
  <button className={`button button--${variant} ${className}`.trim()} {...props}>
    {children}
  </button>
);

function AuthScreen({ onLogin, onSignup, librariesLoaded, message }) {
  const [mode, setMode] = useState('login');
  const [name, setName] = useState('');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');

  const handleSubmit = (event) => {
    event.preventDefault();
    const payload = { name, email, password };

    if (mode === 'login') {
      onLogin(payload);
      return;
    }

    onSignup(payload);
  };

  return (
    <div className="auth-shell">
      <div className="ambient ambient--one" />
      <div className="ambient ambient--two" />
      <div className="auth-grid">
        <section className="glass-card hero-card reveal-up">
          <div className="hero-chip">Interactive Excel Intelligence</div>
          <h1>Make spreadsheets feel alive.</h1>
          <p>
            Upload Excel files, generate animated charts, search raw rows, and keep your analysis history locally in
            the browser with no backend permission errors.
          </p>
          <div className="hero-features">
            <div>
              <IconWrap><UploadIcon /></IconWrap>
              <span>Drag-and-drop uploads</span>
            </div>
            <div>
              <IconWrap><ChartIcon /></IconWrap>
              <span>Bar, line, and area charts</span>
            </div>
            <div>
              <IconWrap><SparkIcon /></IconWrap>
              <span>Insight cards and preview search</span>
            </div>
          </div>
          <div className="hero-stats">
            <div>
              <strong>{librariesLoaded ? 'Ready' : 'Loading'}</strong>
              <span>Excel parser</span>
            </div>
            <div>
              <strong>Local-first</strong>
              <span>Works without Firestore</span>
            </div>
            <div>
              <strong>Animated</strong>
              <span>Modern dashboard feel</span>
            </div>
          </div>
        </section>

        <section className="glass-card auth-card reveal-up reveal-delay-1">
          <div className="auth-card__header">
            <div>
              <p className="eyebrow">Access</p>
              <h2>{mode === 'login' ? 'Sign in to continue' : 'Create your workspace'}</h2>
            </div>
            <div className="segmented-control">
              <button type="button" className={mode === 'login' ? 'active' : ''} onClick={() => setMode('login')}>Login</button>
              <button type="button" className={mode === 'signup' ? 'active' : ''} onClick={() => setMode('signup')}>Sign up</button>
            </div>
          </div>

          <form className="auth-form" onSubmit={handleSubmit}>
            {mode === 'signup' && (
              <label>
                <span>Name</span>
                <input value={name} onChange={(event) => setName(event.target.value)} placeholder="Your name" />
              </label>
            )}
            <label>
              <span>Email</span>
              <input type="email" value={email} onChange={(event) => setEmail(event.target.value)} placeholder="you@example.com" />
            </label>
            <label>
              <span>Password</span>
              <input type="password" value={password} onChange={(event) => setPassword(event.target.value)} placeholder="Minimum 6 characters" />
            </label>
            <Button type="submit">{mode === 'login' ? 'Open Dashboard' : 'Create Account'}</Button>
          </form>

          <div className="auth-note">
            <strong>Status</strong>
            <span>{message || 'Your uploaded files stay inside this browser unless you export them.'}</span>
          </div>
        </section>
      </div>
    </div>
  );
}

function UploadPanel({ onUpload, onLoadDemo, isUploading }) {
  const fileInputRef = useRef(null);
  const [isDragging, setIsDragging] = useState(false);

  const forwardFile = (file) => {
    if (file) {
      onUpload(file);
    }
  };

  return (
    <section className="glass-card panel reveal-up reveal-delay-1">
      <div className="panel-heading">
        <div>
          <p className="eyebrow">Upload</p>
          <h3>Bring in your workbook</h3>
        </div>
        <div className="status-pill">.xls / .xlsx</div>
      </div>

      <div
        className={`dropzone ${isDragging ? 'dropzone--active' : ''}`}
        onClick={() => fileInputRef.current?.click()}
        onDragEnter={(event) => {
          event.preventDefault();
          setIsDragging(true);
        }}
        onDragOver={(event) => {
          event.preventDefault();
          setIsDragging(true);
        }}
        onDragLeave={(event) => {
          event.preventDefault();
          setIsDragging(false);
        }}
        onDrop={(event) => {
          event.preventDefault();
          setIsDragging(false);
          forwardFile(event.dataTransfer.files?.[0]);
        }}
        role="button"
        tabIndex={0}
        onKeyDown={(event) => {
          if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault();
            fileInputRef.current?.click();
          }
        }}
      >
        <div className="dropzone__icon"><UploadIcon /></div>
        <h4>{isUploading ? 'Analyzing workbook...' : 'Drop your Excel file here'}</h4>
        <p>We will parse the first sheet, generate metrics, and save it in your dashboard history.</p>
        <div className="dropzone__actions">
          <Button type="button" variant="secondary">Choose File</Button>
          <Button
            type="button"
            variant="ghost"
            onClick={(event) => {
              event.stopPropagation();
              onLoadDemo();
            }}
          >
            Load Demo Data
          </Button>
        </div>
        <input
          ref={fileInputRef}
          className="sr-only"
          type="file"
          accept=".xls,.xlsx"
          onChange={(event) => forwardFile(event.target.files?.[0])}
        />
      </div>
    </section>
  );
}

function StatsStrip({ files, activeFile }) {
  const cards = [
    { label: 'Uploaded files', value: files.length, tone: 'teal' },
    { label: 'Rows processed', value: files.reduce((sum, file) => sum + file.rowCount, 0), tone: 'orange' },
    { label: 'Active columns', value: activeFile?.headers.length || 0, tone: 'blue' },
    { label: 'Numeric metrics', value: activeFile?.numericHeaders.length || 0, tone: 'rose' },
  ];

  return (
    <section className="stats-strip reveal-up">
      {cards.map((card) => (
        <article key={card.label} className={`metric-card metric-card--${card.tone}`}>
          <span>{card.label}</span>
          <strong>{formatNumber(card.value)}</strong>
        </article>
      ))}
    </section>
  );
}

function HistoryPanel({ files, activeFileId, onSelect, onDelete }) {
  return (
    <section className="glass-card panel reveal-up reveal-delay-2">
      <div className="panel-heading">
        <div>
          <p className="eyebrow">Library</p>
          <h3>Upload history</h3>
        </div>
        <div className="status-pill">{files.length} files</div>
      </div>

      {files.length === 0 ? (
        <div className="empty-state">
          <div className="empty-state__icon"><FileIcon /></div>
          <h4>No files yet</h4>
          <p>Your analyzed workbooks will show up here with time, rows, and columns.</p>
        </div>
      ) : (
        <div className="history-list">
          {files.map((file) => (
            <button
              key={file.id}
              type="button"
              className={`history-item ${activeFileId === file.id ? 'history-item--active' : ''}`}
              onClick={() => onSelect(file.id)}
            >
              <div>
                <strong>{file.name}</strong>
                <span>{formatDate(file.uploadedAt)}</span>
              </div>
              <div className="history-item__meta">
                <span>{file.rowCount} rows</span>
                <span>{file.headers.length} cols</span>
                <Button
                  type="button"
                  variant="ghost"
                  className="history-delete"
                  onClick={(event) => {
                    event.stopPropagation();
                    onDelete(file.id);
                  }}
                >
                  Delete
                </Button>
              </div>
            </button>
          ))}
        </div>
      )}
    </section>
  );
}

function AnalysisPanel({ file, onExportCsv }) {
  const quickPrompts = [
    'What trend do you see?',
    'Which segment is performing best?',
    'What should I chart next?',
    'Where is the biggest risk?',
  ];
  const chartRef = useRef(null);
  const [chartType, setChartType] = useState('bar');
  const [xAxis, setXAxis] = useState(file.headers[0] || '');
  const [yAxis, setYAxis] = useState(file.numericHeaders[0] || '');
  const [search, setSearch] = useState('');
  const [sortConfig, setSortConfig] = useState({ key: '', direction: 'asc' });
  const [aiQuestion, setAiQuestion] = useState('');
  const [aiAnswer, setAiAnswer] = useState('');

  useEffect(() => {
    setChartType('bar');
    setXAxis(file.headers[0] || '');
    setYAxis(file.numericHeaders[0] || '');
    setSearch('');
    setSortConfig({ key: '', direction: 'asc' });
    setAiQuestion('');
    setAiAnswer('');
  }, [file]);

  const chartData = useMemo(() => getChartData(file.rows, xAxis, yAxis), [file.rows, xAxis, yAxis]);
  const insights = useMemo(() => buildInsights(file.rows, file.numericHeaders), [file.rows, file.numericHeaders]);
  const aiCards = useMemo(() => buildAiCards(file), [file]);

  const previewRows = useMemo(() => {
    const query = search.trim().toLowerCase();
    let rows = !query
      ? [...file.rows]
      : file.rows.filter((row) => file.headers.some((header) => String(row[header] ?? '').toLowerCase().includes(query)));

    if (sortConfig.key) {
      rows.sort((left, right) => {
        const leftValue = left[sortConfig.key];
        const rightValue = right[sortConfig.key];
        const leftNumber = Number(leftValue);
        const rightNumber = Number(rightValue);
        const bothNumeric = leftValue !== '' && rightValue !== '' && !Number.isNaN(leftNumber) && !Number.isNaN(rightNumber);

        let result = 0;
        if (bothNumeric) {
          result = leftNumber - rightNumber;
        } else {
          result = String(leftValue ?? '').localeCompare(String(rightValue ?? ''));
        }

        return sortConfig.direction === 'asc' ? result : -result;
      });
    }

    return rows.slice(0, 8);
  }, [file.headers, file.rows, search, sortConfig]);

  const handleAskAi = (question) => {
    const nextQuestion = question.trim();
    if (!nextQuestion) {
      return;
    }

    setAiQuestion(nextQuestion);
    setAiAnswer(generateAiResponse(file, nextQuestion));
  };

  const handleSort = (header) => {
    setSortConfig((current) => ({
      key: header,
      direction: current.key === header && current.direction === 'asc' ? 'desc' : 'asc',
    }));
  };

  const downloadChart = async () => {
    if (!chartRef.current || !window.html2canvas) {
      return;
    }

    const canvas = await window.html2canvas(chartRef.current, { backgroundColor: '#07111c', scale: 2 });
    const link = document.createElement('a');
    link.download = `${file.name.replace(/\.[^.]+$/, '') || 'analysis'}-chart.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
  };

  const renderChart = () => {
    const sharedProps = {
      data: chartData,
      margin: { top: 10, right: 10, left: 0, bottom: 0 },
    };

    if (chartType === 'line') {
      return (
        <LineChart {...sharedProps}>
          <CartesianGrid stroke="rgba(148, 163, 184, 0.18)" strokeDasharray="3 3" />
          <XAxis dataKey="name" stroke="#94a3b8" tickLine={false} axisLine={false} />
          <YAxis stroke="#94a3b8" tickLine={false} axisLine={false} />
          <Tooltip contentStyle={{ background: '#0f172a', border: '1px solid rgba(148,163,184,0.2)', borderRadius: 16 }} />
          <Legend />
          <Line type="monotone" dataKey="value" stroke="#f97316" strokeWidth={3} dot={{ r: 4 }} isAnimationActive />
        </LineChart>
      );
    }

    if (chartType === 'area') {
      return (
        <AreaChart {...sharedProps}>
          <defs>
            <linearGradient id="analysisFill" x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%" stopColor="#2563eb" stopOpacity={0.85} />
              <stop offset="95%" stopColor="#2563eb" stopOpacity={0.05} />
            </linearGradient>
          </defs>
          <CartesianGrid stroke="rgba(148, 163, 184, 0.18)" strokeDasharray="3 3" />
          <XAxis dataKey="name" stroke="#94a3b8" tickLine={false} axisLine={false} />
          <YAxis stroke="#94a3b8" tickLine={false} axisLine={false} />
          <Tooltip contentStyle={{ background: '#0f172a', border: '1px solid rgba(148,163,184,0.2)', borderRadius: 16 }} />
          <Legend />
          <Area type="monotone" dataKey="value" stroke="#2563eb" fill="url(#analysisFill)" strokeWidth={3} isAnimationActive />
        </AreaChart>
      );
    }

    return (
      <BarChart {...sharedProps}>
        <CartesianGrid stroke="rgba(148, 163, 184, 0.18)" strokeDasharray="3 3" />
        <XAxis dataKey="name" stroke="#94a3b8" tickLine={false} axisLine={false} />
        <YAxis stroke="#94a3b8" tickLine={false} axisLine={false} />
        <Tooltip contentStyle={{ background: '#0f172a', border: '1px solid rgba(148,163,184,0.2)', borderRadius: 16 }} />
        <Legend />
        <Bar dataKey="value" radius={[12, 12, 4, 4]} isAnimationActive>
          {chartData.map((entry, index) => (
            <Cell key={`${entry.name}-${index}`} fill={CHART_COLORS[index % CHART_COLORS.length]} />
          ))}
        </Bar>
      </BarChart>
    );
  };

  return (
    <section className="analysis-stack">
      <article className="glass-card panel reveal-up">
        <div className="panel-heading">
          <div>
            <p className="eyebrow">Analysis</p>
            <h3>{file.name}</h3>
          </div>
          <div className="status-pill">{file.rowCount} rows</div>
        </div>

        <div className="overview-banner">
          <div>
            <span>Sheet</span>
            <strong>{file.sheetName}</strong>
          </div>
          <div>
            <span>Columns</span>
            <strong>{file.headers.length}</strong>
          </div>
          <div>
            <span>Numeric</span>
            <strong>{file.numericHeaders.length}</strong>
          </div>
        </div>

        <div className="insight-grid">
          {insights.length === 0 ? (
            <div className="empty-inline">This file needs at least one numeric column to generate insight cards.</div>
          ) : (
            insights.map((insight) => (
              <div key={insight.label} className="insight-card">
                <span>{insight.label}</span>
                <strong>{formatNumber(insight.total)}</strong>
                <small>Avg {formatNumber(insight.average)} • Peak {formatNumber(insight.peak)}</small>
              </div>
            ))
          )}
        </div>
      </article>

      <article className="glass-card panel reveal-up reveal-delay-1">
        <div className="panel-heading">
          <div>
            <p className="eyebrow">AI Analyst</p>
            <h3>Ask for a quick read</h3>
          </div>
          <div className="status-pill">Local AI</div>
        </div>

        <div className="ai-card-grid">
          {aiCards.length === 0 ? (
            <div className="empty-inline">Upload a dataset with numeric values to unlock AI-style analysis.</div>
          ) : (
            aiCards.map((card) => (
              <div key={card.title} className="ai-mini-card">
                <span>{card.title}</span>
                <strong>{card.body}</strong>
              </div>
            ))
          )}
        </div>

        <div className="ai-prompt-row">
          {quickPrompts.map((prompt) => (
            <Button key={prompt} type="button" variant="ghost" className="prompt-chip" onClick={() => handleAskAi(prompt)}>
              {prompt}
            </Button>
          ))}
        </div>

        <div className="ai-query">
          <input
            className="search-input"
            value={aiQuestion}
            onChange={(event) => setAiQuestion(event.target.value)}
            placeholder="Ask about trends, top segments, risks, or chart ideas"
          />
          <Button type="button" onClick={() => handleAskAi(aiQuestion)}>Analyze</Button>
        </div>

        {aiAnswer && <div className="ai-answer">{aiAnswer}</div>}
      </article>

      <article className="glass-card panel reveal-up reveal-delay-1">
        <div className="panel-heading">
          <div>
            <p className="eyebrow">Visualize</p>
            <h3>Interactive chart studio</h3>
          </div>
          <div className="action-row">
            <Button type="button" variant="ghost" onClick={() => onExportCsv(file)}>Export CSV</Button>
            <Button type="button" variant="secondary" onClick={downloadChart}>Download PNG</Button>
          </div>
        </div>

        <div className="toolbar-grid">
          <label>
            <span>Chart type</span>
            <select value={chartType} onChange={(event) => setChartType(event.target.value)}>
              <option value="bar">Bar</option>
              <option value="line">Line</option>
              <option value="area">Area</option>
            </select>
          </label>
          <label>
            <span>X-axis</span>
            <select value={xAxis} onChange={(event) => setXAxis(event.target.value)}>
              {file.headers.map((header) => (
                <option key={header} value={header}>{header}</option>
              ))}
            </select>
          </label>
          <label>
            <span>Y-axis</span>
            <select value={yAxis} onChange={(event) => setYAxis(event.target.value)}>
              {file.numericHeaders.map((header) => (
                <option key={header} value={header}>{header}</option>
              ))}
            </select>
          </label>
        </div>

        <div className="chart-shell" ref={chartRef}>
          {chartData.length === 0 ? (
            <div className="empty-inline">Choose one text column and one numeric column to draw the chart.</div>
          ) : (
            <ResponsiveContainer width="100%" height={340}>{renderChart()}</ResponsiveContainer>
          )}
        </div>
      </article>

      <article className="glass-card panel reveal-up reveal-delay-2">
        <div className="panel-heading">
          <div>
            <p className="eyebrow">Preview</p>
            <h3>Search and sort your rows</h3>
          </div>
          <input className="search-input" value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search any value" />
        </div>

        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                {file.headers.map((header) => (
                  <th key={header}>
                    <button type="button" className="table-sort" onClick={() => handleSort(header)}>
                      {header}
                      {sortConfig.key === header ? (sortConfig.direction === 'asc' ? ' ↑' : ' ↓') : ''}
                    </button>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {previewRows.length === 0 ? (
                <tr>
                  <td colSpan={file.headers.length}>No matching rows for that search.</td>
                </tr>
              ) : (
                previewRows.map((row) => (
                  <tr key={row.__rowId}>
                    {file.headers.map((header) => (
                      <td key={`${row.__rowId}-${header}`}>{String(row[header] ?? '')}</td>
                    ))}
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </article>
    </section>
  );
}

export default function App() {
  const [librariesLoaded, setLibrariesLoaded] = useState(false);
  const [notification, setNotification] = useState({ type: 'info', message: '' });
  const [currentUser, setCurrentUser] = useState(() => readJson(SESSION_KEY, null));
  const [files, setFiles] = useState(() => []);
  const [activeFileId, setActiveFileId] = useState(null);
  const [isUploading, setIsUploading] = useState(false);

  useEffect(() => {
    const loadScript = (src, id) =>
      new Promise((resolve, reject) => {
        if (document.getElementById(id)) {
          resolve();
          return;
        }

        const script = document.createElement('script');
        script.id = id;
        script.src = src;
        script.async = true;
        script.onload = () => resolve();
        script.onerror = () => reject(new Error(`Unable to load ${src}`));
        document.head.appendChild(script);
      });

    Promise.all([
      loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'xlsx-script'),
      loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js', 'html2canvas-script'),
    ])
      .then(() => {
        setLibrariesLoaded(true);
        setNotification({ type: 'success', message: 'Dashboard engine is ready.' });
      })
      .catch((error) => {
        console.error(error);
        setNotification({ type: 'error', message: 'Required browser libraries failed to load. Refresh and try again.' });
      });
  }, []);

  useEffect(() => {
    if (!currentUser?.email) {
      setFiles([]);
      setActiveFileId(null);
      return;
    }

    const storedFiles = loadFiles(currentUser.email);
    setFiles(storedFiles);
    setActiveFileId(storedFiles[0]?.id || null);
  }, [currentUser]);

  useEffect(() => {
    if (!notification.message) {
      return undefined;
    }

    const timeout = window.setTimeout(() => setNotification({ type: 'info', message: '' }), 3200);
    return () => window.clearTimeout(timeout);
  }, [notification]);

  const activeFile = useMemo(() => files.find((file) => file.id === activeFileId) || files[0] || null, [files, activeFileId]);

  const persistFiles = (nextFiles, nextActiveId) => {
    setFiles(nextFiles);
    setActiveFileId(nextActiveId);
    saveFiles(currentUser.email, nextFiles);
  };

  const createStoredFile = ({ name, sheetName, headers, rows, size = 0 }) => ({
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    name,
    sheetName,
    uploadedAt: new Date().toISOString(),
    headers,
    rows,
    rowCount: rows.length,
    numericHeaders: getNumericHeaders(headers, rows),
    size,
  });

  const handleSignup = ({ name, email, password }) => {
    const normalizedEmail = email.trim().toLowerCase();
    const cleanName = name.trim();

    if (!cleanName || !normalizedEmail || password.length < 6) {
      setNotification({ type: 'error', message: 'Name, email, and a 6+ character password are required.' });
      return;
    }

    const users = readJson(USERS_KEY, []);
    if (users.some((user) => user.email === normalizedEmail)) {
      setNotification({ type: 'error', message: 'That account already exists. Try logging in instead.' });
      return;
    }

    const nextUser = {
      id: normalizedEmail.replace(/[^a-z0-9]+/g, '-'),
      name: cleanName,
      email: normalizedEmail,
      password,
      createdAt: new Date().toISOString(),
    };

    writeJson(USERS_KEY, [...users, nextUser]);
    writeJson(SESSION_KEY, nextUser);
    setCurrentUser(nextUser);
    setNotification({ type: 'success', message: 'Account created successfully.' });
  };

  const handleLogin = ({ email, password }) => {
    const normalizedEmail = email.trim().toLowerCase();
    const users = readJson(USERS_KEY, []);
    const matchedUser = users.find((user) => user.email === normalizedEmail && user.password === password);

    if (!matchedUser) {
      setNotification({ type: 'error', message: 'Incorrect email or password.' });
      return;
    }

    writeJson(SESSION_KEY, matchedUser);
    setCurrentUser(matchedUser);
    setNotification({ type: 'success', message: `Welcome back, ${matchedUser.name}.` });
  };

  const handleLogout = () => {
    window.localStorage.removeItem(SESSION_KEY);
    setCurrentUser(null);
    setNotification({ type: 'success', message: 'Signed out successfully.' });
  };

  const handleUpload = (file) => {
    if (!librariesLoaded || !window.XLSX) {
      setNotification({ type: 'error', message: 'The Excel library is still loading. Please wait a moment.' });
      return;
    }

    if (!file || !/\.(xlsx|xls)$/i.test(file.name)) {
      setNotification({ type: 'error', message: 'Please upload a valid .xls or .xlsx file.' });
      return;
    }

    setIsUploading(true);
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const workbook = window.XLSX.read(new Uint8Array(event.target.result), { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const sheetRows = window.XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });
        const { headers, rows } = toTableData(sheetRows);

        if (!headers.length || !rows.length) {
          setNotification({ type: 'error', message: 'This workbook is empty. Add data and try again.' });
          setIsUploading(false);
          return;
        }

        const nextFile = createStoredFile({
          name: file.name,
          sheetName: firstSheetName,
          headers,
          rows,
          size: file.size,
        });

        const nextFiles = [nextFile, ...files];
        persistFiles(nextFiles, nextFile.id);
        setNotification({ type: 'success', message: `${file.name} uploaded and analyzed successfully.` });
      } catch (error) {
        console.error(error);
        setNotification({ type: 'error', message: 'We could not parse that Excel file.' });
      } finally {
        setIsUploading(false);
      }
    };

    reader.onerror = () => {
      setIsUploading(false);
      setNotification({ type: 'error', message: 'There was a problem reading that file.' });
    };

    reader.readAsArrayBuffer(file);
  };

  const handleLoadDemo = () => {
    const demo = buildDemoWorkbook();
    const { headers, rows } = toTableData(demo.rows);
    const nextFile = createStoredFile({
      name: demo.name,
      sheetName: demo.sheetName,
      headers,
      rows,
    });
    const nextFiles = [nextFile, ...files];
    persistFiles(nextFiles, nextFile.id);
    setNotification({ type: 'success', message: 'Demo dataset loaded into your dashboard.' });
  };

  const handleDeleteFile = (fileId) => {
    const nextFiles = files.filter((file) => file.id !== fileId);
    persistFiles(nextFiles, nextFiles[0]?.id || null);
    setNotification({ type: 'success', message: 'File removed from your history.' });
  };

  const handleExportCsv = (file) => {
    const lines = [file.headers.join(',')].concat(
      file.rows.map((row) =>
        file.headers
          .map((header) => {
            const value = String(row[header] ?? '');
            return /[",\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
          })
          .join(',')
      )
    );

    downloadTextFile(`${file.name.replace(/\.[^.]+$/, '') || 'dataset'}.csv`, lines.join('\n'), 'text/csv;charset=utf-8;');
    setNotification({ type: 'success', message: 'CSV export downloaded.' });
  };

  if (!currentUser) {
    return (
      <AuthScreen
        onLogin={handleLogin}
        onSignup={handleSignup}
        librariesLoaded={librariesLoaded}
        message={notification.message}
      />
    );
  }

  return (
    <div className="app-shell">
      <div className="ambient ambient--one" />
      <div className="ambient ambient--two" />

      {notification.message && <div className={`toast toast--${notification.type}`}>{notification.message}</div>}

      <header className="topbar reveal-up">
        <div>
          <p className="eyebrow">Excel Analytics Platform</p>
          <h1>{currentUser.name}'s dashboard</h1>
        </div>
        <div className="topbar__actions">
          <div className="user-pill">{currentUser.email}</div>
          <Button type="button" variant="secondary" onClick={handleLogout}>
            <span className="button__icon"><LogoutIcon /></span>
            Logout
          </Button>
        </div>
      </header>

      <main className="dashboard-grid">
        <section className="dashboard-main">
          <section className="glass-card spotlight reveal-up">
            <div>
              <p className="eyebrow">Overview</p>
              <h2>Turn spreadsheets into a clear visual story.</h2>
              <p>
                Upload a workbook, switch chart modes, inspect key metrics, export CSV, and search or sort your rows without leaving the page.
              </p>
            </div>
            <div className="spotlight__chips">
              <span>Local storage</span>
              <span>Animated charts</span>
              <span>Demo dataset</span>
            </div>
          </section>

          <StatsStrip files={files} activeFile={activeFile} />
          <UploadPanel onUpload={handleUpload} onLoadDemo={handleLoadDemo} isUploading={isUploading} />
          {activeFile && <AnalysisPanel file={activeFile} onExportCsv={handleExportCsv} />}
        </section>

        <aside className="dashboard-side">
          <HistoryPanel files={files} activeFileId={activeFile?.id} onSelect={setActiveFileId} onDelete={handleDeleteFile} />
        </aside>
      </main>
    </div>
  );
}
