import React, { useState } from 'react';
import {
  LayoutDashboard,
  FileText,
  CalendarCheck,
  FileSpreadsheet,
  AlertTriangle,
  TrendingUp,
  FolderArchive,
  Users,
  LogOut,
  ShieldCheck,
  Search,
  Menu,
  X,
  ExternalLink,
  Database,
  Zap,
  CheckCircle2,
  Lock,
  Building2,
  HardHat,
  ChevronRight,
  RefreshCw
} from 'lucide-react';
import { HSE_UTILITY_LINKS, SIDEBAR_NAV_ITEMS } from '../data/hseUtilities';
import { HSEUtilityLink, UserSession } from '../types';

interface DashboardLayoutProps {
  userSession?: UserSession;
  activeNavId?: string;
  onLogout?: () => void;
}

export const DashboardLayout: React.FC<DashboardLayoutProps> = ({
  userSession = {
    username: 'admin',
    role: 'admin',
    projectCode: 'BI-10-10303',
    facility: 'Abqaiq Production Facilities',
  },
  activeNavId = 'dashboard',
  onLogout,
}) => {
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedFormat, setSelectedFormat] = useState<string>('all');

  // Map icon names to Lucide components
  const renderIcon = (iconName: string, className: string = 'w-5 h-5') => {
    switch (iconName) {
      case 'LayoutDashboard': return <LayoutDashboard className={className} />;
      case 'FileText': return <FileText className={className} />;
      case 'CalendarCheck': return <CalendarCheck className={className} />;
      case 'FileSpreadsheet': return <FileSpreadsheet className={className} />;
      case 'AlertTriangle': return <AlertTriangle className={className} />;
      case 'TrendingUp': return <TrendingUp className={className} />;
      case 'FolderArchive': return <FolderArchive className={className} />;
      case 'Users': return <Users className={className} />;
      case 'ShieldCheck': return <ShieldCheck className={className} />;
      default: return <FileText className={className} />;
    }
  };

  const filteredUtilities = HSE_UTILITY_LINKS.filter((item) => {
    // Role filter
    if (item.adminOnly && userSession.role !== 'admin') {
      return false;
    }
    // Format filter
    if (selectedFormat !== 'all' && item.format !== selectedFormat) {
      return false;
    }
    // Search query
    if (searchQuery.trim() !== '') {
      const q = searchQuery.toLowerCase();
      return (
        item.title.toLowerCase().includes(q) ||
        item.description.toLowerCase().includes(q) ||
        item.tags.some((tag) => tag.toLowerCase().includes(q))
      );
    }
    return true;
  });

  const getCardBorderColor = (theme: HSEUtilityLink['themeColor']) => {
    switch (theme) {
      case 'green': return 'border-t-[#009966] hover:border-t-[#007a52]';
      case 'teal': return 'border-t-[#00b3b3] hover:border-t-[#008f8f]';
      case 'amber': return 'border-t-[#ffcc00] hover:border-t-[#d4a800]';
      case 'navy': return 'border-t-[#002244] hover:border-t-[#001730]';
    }
  };

  const getButtonBg = (theme: HSEUtilityLink['themeColor']) => {
    switch (theme) {
      case 'green': return 'bg-[#009966] hover:bg-[#008055] text-white shadow-emerald-900/20';
      case 'teal': return 'bg-[#00b3b3] hover:bg-[#009999] text-white shadow-teal-900/20';
      case 'amber': return 'bg-[#002244] hover:bg-[#001730] text-[#ffcc00] shadow-blue-950/20';
      case 'navy': return 'bg-[#002244] hover:bg-[#001730] text-white shadow-blue-950/20';
    }
  };

  const handleLogout = () => {
    if (onLogout) {
      onLogout();
    } else {
      window.location.href = '/logout';
    }
  };

  return (
    <div id="dashboardRoot" className="min-h-screen bg-slate-900 text-slate-100 flex flex-col font-sans selection:bg-[#ffcc00] selection:text-[#002244]">
      {/* Mobile Sidebar Overlay */}
      {sidebarOpen && (
        <div
          id="sidebarBackdrop"
          className="fixed inset-0 bg-black/60 z-40 lg:hidden backdrop-blur-sm transition-opacity"
          onClick={() => setSidebarOpen(false)}
        />
      )}

      <div className="flex flex-1 w-full overflow-hidden">
        {/* ========================================================================= */}
        {/* SIDEBAR COMPONENT                                                        */}
        {/* ========================================================================= */}
        <aside
          id="dashboardSidebar"
          className={`fixed inset-y-0 left-0 z-50 w-72 bg-[#001a33] border-r border-slate-800 flex flex-col justify-between transition-transform duration-300 ease-in-out lg:static lg:translate-x-0 ${
            sidebarOpen ? 'translate-x-0' : '-translate-x-full'
          }`}
        >
          {/* Sidebar Header */}
          <div className="flex flex-col border-b border-slate-800/80 bg-[#001428] px-5 py-4">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-[#009966] via-[#00b3b3] to-[#ffcc00] p-[2px] shadow-lg flex items-center justify-center">
                  <div className="w-full h-full bg-[#002244] rounded-[10px] flex items-center justify-center">
                    <ShieldCheck className="w-5 h-5 text-[#ffcc00]" />
                  </div>
                </div>
                <div>
                  <h1 className="text-sm font-bold tracking-wider text-white uppercase flex items-center gap-1.5">
                    Central HSE <span className="text-[#ffcc00]">Ops</span>
                  </h1>
                  <p className="text-[11px] font-medium text-slate-400">BI-10-10303 Abqaiq</p>
                </div>
              </div>

              {/* Close button on mobile */}
              <button
                id="btnCloseSidebar"
                className="lg:hidden p-1.5 text-slate-400 hover:text-white rounded-lg hover:bg-slate-800"
                onClick={() => setSidebarOpen(false)}
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Corporate Partners Pills */}
            <div className="mt-3 flex items-center gap-1.5 text-[10px] font-semibold text-slate-400 uppercase tracking-wider">
              <span className="bg-slate-800/90 text-slate-300 px-2 py-0.5 rounded border border-slate-700/60">HEISCO</span>
              <span className="text-slate-600">&bull;</span>
              <span className="bg-slate-800/90 text-slate-300 px-2 py-0.5 rounded border border-slate-700/60">ENPPI</span>
              <span className="text-slate-600">&bull;</span>
              <span className="bg-emerald-950/80 text-emerald-300 px-2 py-0.5 rounded border border-emerald-800/50">ARAMCO</span>
            </div>
          </div>

          {/* Navigation Links Area */}
          <div className="flex-1 overflow-y-auto px-3 py-4 space-y-1 scrollbar-thin scrollbar-thumb-slate-700">
            <div className="px-3 pb-2 text-[10px] font-bold uppercase tracking-wider text-slate-400">
              Operations &amp; Reporting
            </div>

            {SIDEBAR_NAV_ITEMS.map((item) => {
              if (item.adminOnly && userSession.role !== 'admin') {
                return null;
              }
              const isActive = activeNavId === item.id;
              return (
                <a
                  key={item.id}
                  id={`sideNav-${item.id}`}
                  href={item.href}
                  className={`group flex items-center justify-between px-3 py-2.5 rounded-xl text-xs font-semibold transition-all ${
                    isActive
                      ? 'bg-gradient-to-r from-[#002b4d] to-[#003866] text-[#ffcc00] shadow-md border border-[#ffcc00]/30'
                      : 'text-slate-300 hover:text-white hover:bg-slate-800/60'
                  }`}
                >
                  <div className="flex items-center gap-3">
                    <span className={`transition-colors ${isActive ? 'text-[#ffcc00]' : 'text-slate-400 group-hover:text-white'}`}>
                      {renderIcon(item.icon, 'w-4 h-4')}
                    </span>
                    <span>{item.label}</span>
                  </div>

                  {item.badge && (
                    <span className="text-[10px] bg-slate-800/90 text-slate-300 px-1.5 py-0.5 rounded font-mono border border-slate-700">
                      {item.badge}
                    </span>
                  )}
                  {isActive && !item.badge && (
                    <span className="w-1.5 h-1.5 rounded-full bg-[#ffcc00]" />
                  )}
                </a>
              );
            })}

            <div className="pt-4 px-3 pb-2 text-[10px] font-bold uppercase tracking-wider text-slate-400">
              Session &amp; Security
            </div>

            <a
              id="sideNavResetPassword"
              href="/reset_password"
              className="group flex items-center gap-3 px-3 py-2 rounded-xl text-xs font-medium text-slate-400 hover:text-white hover:bg-slate-800/60 transition-colors"
            >
              <Lock className="w-4 h-4 text-slate-400 group-hover:text-white" />
              <span>Change Password</span>
            </a>
          </div>

          {/* Sidebar Footer / System Badge */}
          <div className="p-3 border-t border-slate-800 bg-[#001428]/90 space-y-2">
            <div className="bg-slate-900/90 border border-slate-800 rounded-lg p-2.5 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Database className="w-3.5 h-3.5 text-emerald-400" />
                <span className="text-[11px] text-slate-300 font-medium">SQLite users.db</span>
              </div>
              <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-emerald-400">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse" />
                Active
              </span>
            </div>

            {/* Quick user profile info */}
            <div className="flex items-center justify-between px-1 pt-1">
              <div className="flex items-center gap-2">
                <div className="w-7 h-7 rounded-full bg-[#ffcc00] text-[#002244] font-bold flex items-center justify-center text-xs shadow-inner">
                  {userSession.username.slice(0, 2).toUpperCase()}
                </div>
                <div>
                  <p className="text-xs font-semibold text-white leading-tight">{userSession.username}</p>
                  <p className="text-[10px] text-[#ffcc00] uppercase font-bold tracking-wide">{userSession.role}</p>
                </div>
              </div>

              <button
                id="sidebarLogoutBtn"
                onClick={handleLogout}
                title="Sign Out"
                className="p-1.5 text-rose-400 hover:text-white hover:bg-rose-900/60 rounded-lg transition-colors"
              >
                <LogOut className="w-4 h-4" />
              </button>
            </div>
          </div>
        </aside>

        {/* ========================================================================= */}
        {/* MAIN APPLICATION CONTAINER (Top Nav + Main Content)                     */}
        {/* ========================================================================= */}
        <div className="flex-1 flex flex-col min-w-0 bg-[#002244] text-slate-100 overflow-y-auto">
          
          {/* ======================================================================= */}
          {/* TOP NAVIGATION BAR                                                      */}
          {/* ======================================================================= */}
          <header
            id="topNavBar"
            className="sticky top-0 z-30 bg-[#001a33]/95 backdrop-blur border-b-2 border-[#ffcc00] px-4 lg:px-8 py-3 flex items-center justify-between gap-4 shadow-xl"
          >
            {/* Left: Mobile Toggle & Breadcrumbs */}
            <div className="flex items-center gap-3">
              <button
                id="btnToggleSidebar"
                className="lg:hidden p-2 rounded-lg bg-slate-800 text-slate-300 hover:text-white hover:bg-slate-700 transition"
                onClick={() => setSidebarOpen(!sidebarOpen)}
                aria-label="Toggle navigation menu"
              >
                <Menu className="w-5 h-5" />
              </button>

              <div>
                <div className="flex items-center gap-2 text-xs font-semibold text-slate-300">
                  <span className="text-[#ffcc00] font-bold tracking-wider uppercase">Central HSE Operations</span>
                  <span className="text-slate-500">/</span>
                  <span className="text-white">Utilities Hub</span>
                </div>
                <div className="hidden sm:flex items-center gap-2 text-[11px] text-slate-400 mt-0.5">
                  <span className="inline-flex items-center gap-1 font-mono text-emerald-400">
                    <CheckCircle2 className="w-3 h-3 text-emerald-400" /> Cost Center 43300301
                  </span>
                  <span>&bull;</span>
                  <span>{userSession.facility}</span>
                </div>
              </div>
            </div>

            {/* Center: Search / Filter input */}
            <div className="hidden md:flex flex-1 max-w-xs relative">
              <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                id="topNavSearchInput"
                type="text"
                placeholder="Quick search HSE utilities..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-9 pr-3 py-1.5 text-xs bg-slate-800/80 border border-slate-700/80 rounded-lg text-white placeholder-slate-400 focus:outline-none focus:border-[#ffcc00] focus:ring-1 focus:ring-[#ffcc00] transition"
              />
              {searchQuery && (
                <button
                  onClick={() => setSearchQuery('')}
                  className="absolute right-2.5 top-1/2 -translate-y-1/2 text-xs text-slate-400 hover:text-white"
                >
                  &times;
                </button>
              )}
            </div>

            {/* Right: User Details, Role Badge & Logout Button */}
            <div className="flex items-center gap-3">
              {/* Role badge */}
              <div className="hidden sm:flex items-center gap-2 bg-[#ffcc00] text-[#002244] font-bold text-xs px-3 py-1 rounded-full shadow-sm">
                <HardHat className="w-3.5 h-3.5 text-[#002244]" />
                <span className="tracking-wider text-[11px]">{userSession.role.toUpperCase()}</span>
              </div>

              {/* User name indicator */}
              <div className="hidden sm:block text-right">
                <span className="block text-xs font-semibold text-white leading-tight">
                  {userSession.username}
                </span>
                <span className="block text-[10px] text-slate-400">Authenticated</span>
              </div>

              {/* Top Navigation Sign Out Button */}
              <button
                id="topNavLogoutBtn"
                onClick={handleLogout}
                className="flex items-center gap-2 bg-[#b91c1c] hover:bg-[#dc2626] active:scale-95 text-white font-bold text-xs px-3.5 py-1.5 rounded-lg shadow-md transition duration-150 border border-red-500/40"
              >
                <LogOut className="w-3.5 h-3.5" />
                <span>Sign Out</span>
              </button>
            </div>
          </header>

          {/* ======================================================================= */}
          {/* MAIN CONTENT AREA: HSE UTILITY LINKS & METRICS                          */}
          {/* ======================================================================= */}
          <main id="mainDashboardContent" className="flex-1 p-4 lg:p-8 space-y-8 max-w-7xl w-full mx-auto">
            
            {/* Corporate Branding & Operational Banner */}
            <div className="relative overflow-hidden rounded-2xl bg-gradient-to-r from-[#001730] via-[#002244] to-[#003866] border border-slate-700/80 p-6 lg:p-8 shadow-2xl">
              <div className="relative z-10 flex flex-col md:flex-row md:items-center justify-between gap-6">
                <div>
                  <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-[#ffcc00]/10 border border-[#ffcc00]/30 text-[#ffcc00] text-xs font-bold uppercase tracking-wider mb-3">
                    <Building2 className="w-3.5 h-3.5" /> Saudi Aramco Certified Automation Suite
                  </div>
                  <h2 className="text-2xl lg:text-3xl font-extrabold text-white tracking-tight">
                    HEISCO – ENPPI – SAUDI ARAMCO
                  </h2>
                  <p className="text-sm font-medium text-slate-300 mt-1 max-w-2xl">
                    Abqaiq Production Facilities (BI-10-10303) Central Health, Safety &amp; Environment Operations Console.
                    Directly launch Word, Excel, and PowerPoint automation engines below.
                  </p>
                </div>

                {/* Corporate Logos Display */}
                <div className="flex items-center gap-3 bg-white/95 p-3 rounded-xl shadow-lg border border-white/20 self-start md:self-center">
                  <img src="/static/logos/enppi.png" alt="ENPPI Logo" className="h-9 object-contain" />
                  <div className="w-[1px] h-7 bg-slate-300" />
                  <img src="/static/logos/heisco.png" alt="HEISCO Logo" className="h-9 object-contain" />
                  <div className="w-[1px] h-7 bg-slate-300" />
                  <img src="/static/logos/aramco.png" alt="Saudi Aramco Logo" className="h-9 object-contain" />
                </div>
              </div>

              {/* Decorative background grid element */}
              <div className="absolute right-0 top-0 bottom-0 w-1/3 bg-radial from-[#00b3b3]/10 to-transparent pointer-events-none" />
            </div>

            {/* Quick Status Bar / Metric Cards */}
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="bg-[#001730]/90 border border-slate-800 rounded-xl p-4 flex items-center gap-3.5">
                <div className="w-10 h-10 rounded-lg bg-emerald-500/10 border border-emerald-500/30 flex items-center justify-center text-emerald-400">
                  <FileText className="w-5 h-5" />
                </div>
                <div>
                  <div className="text-xs text-slate-400 font-medium">Observation Reports</div>
                  <div className="text-base font-bold text-white">Landscape .docx</div>
                </div>
              </div>

              <div className="bg-[#001730]/90 border border-slate-800 rounded-xl p-4 flex items-center gap-3.5">
                <div className="w-10 h-10 rounded-lg bg-cyan-500/10 border border-cyan-500/30 flex items-center justify-center text-cyan-400">
                  <CalendarCheck className="w-5 h-5" />
                </div>
                <div>
                  <div className="text-xs text-slate-400 font-medium">Weekly Activity</div>
                  <div className="text-base font-bold text-white">KPIs &amp; Permits</div>
                </div>
              </div>

              <div className="bg-[#001730]/90 border border-slate-800 rounded-xl p-4 flex items-center gap-3.5">
                <div className="w-10 h-10 rounded-lg bg-amber-500/10 border border-amber-500/30 flex items-center justify-center text-[#ffcc00]">
                  <FileSpreadsheet className="w-5 h-5" />
                </div>
                <div>
                  <div className="text-xs text-slate-400 font-medium">JSL Action Register</div>
                  <div className="text-base font-bold text-white">Yellow Block .xlsx</div>
                </div>
              </div>

              <div className="bg-[#001730]/90 border border-slate-800 rounded-xl p-4 flex items-center gap-3.5">
                <div className="w-10 h-10 rounded-lg bg-violet-500/10 border border-violet-500/30 flex items-center justify-center text-violet-400">
                  <AlertTriangle className="w-5 h-5" />
                </div>
                <div>
                  <div className="text-xs text-slate-400 font-medium">Safety Signage</div>
                  <div className="text-base font-bold text-white">4-Language .pptx</div>
                </div>
              </div>
            </div>

            {/* Category Filter Tabs & Utility Header */}
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-800 pb-3">
              <div>
                <h3 className="text-lg font-bold text-white flex items-center gap-2">
                  <LayoutDashboard className="w-5 h-5 text-[#ffcc00]" />
                  <span>HSE Operational Utility Engines</span>
                </h3>
                <p className="text-xs text-slate-400 mt-0.5">
                  Select an operational utility link below to process raw Excel observations or generate formal reports.
                </p>
              </div>

              {/* Format Filter Chips */}
              <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0">
                {[
                  { id: 'all', label: 'All Modules' },
                  { id: 'Word (.docx)', label: 'Word (.docx)' },
                  { id: 'Excel (.xlsx)', label: 'Excel (.xlsx)' },
                  { id: 'PowerPoint (.pptx)', label: 'PowerPoint (.pptx)' },
                  { id: 'Analytics', label: 'KPI Charts' },
                ].map((tab) => (
                  <button
                    key={tab.id}
                    onClick={() => setSelectedFormat(tab.id)}
                    className={`px-3 py-1 rounded-lg text-xs font-semibold whitespace-nowrap transition ${
                      selectedFormat === tab.id
                        ? 'bg-[#ffcc00] text-[#002244] shadow'
                        : 'bg-slate-800 text-slate-300 hover:bg-slate-700 hover:text-white'
                    }`}
                  >
                    {tab.label}
                  </button>
                ))}
              </div>
            </div>

            {/* =================================================================== */}
            {/* HSE UTILITY LINKS GRID                                              */}
            {/* =================================================================== */}
            <div id="hseUtilityGrid" className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
              {filteredUtilities.map((item) => (
                <div
                  key={item.id}
                  id={`card-${item.id}`}
                  className={`bg-white text-slate-900 rounded-2xl p-6 shadow-xl border-t-[5px] flex flex-col justify-between transition-all duration-200 hover:-translate-y-1.5 hover:shadow-2xl ${getCardBorderColor(
                    item.themeColor
                  )}`}
                >
                  <div>
                    {/* Header with Icon & Format Badge */}
                    <div className="flex items-start justify-between gap-3 mb-4">
                      <div className="w-12 h-12 rounded-xl bg-slate-100 flex items-center justify-center text-[#002244] shadow-sm">
                        {renderIcon(item.icon, 'w-6 h-6')}
                      </div>
                      <span className="text-[11px] font-bold uppercase tracking-wider px-2.5 py-1 rounded-md bg-slate-100 text-slate-700 border border-slate-200">
                        {item.badge}
                      </span>
                    </div>

                    {/* Title & Description */}
                    <h4 className="text-lg font-bold text-[#002244] leading-snug mb-2">
                      {item.title}
                    </h4>
                    <p className="text-xs text-slate-600 leading-relaxed mb-4">
                      {item.description}
                    </p>

                    {/* Tag Chips */}
                    <div className="flex flex-wrap gap-1.5 mb-6">
                      {item.tags.map((tag, idx) => (
                        <span
                          key={idx}
                          className="text-[10px] font-medium bg-slate-100 text-slate-700 px-2 py-0.5 rounded border border-slate-200"
                        >
                          {tag}
                        </span>
                      ))}
                    </div>
                  </div>

                  {/* Launch Action Button */}
                  <a
                    id={`btnLaunch-${item.id}`}
                    href={item.href}
                    className={`w-full py-2.5 px-4 rounded-xl font-bold text-xs tracking-wide text-center flex items-center justify-center gap-2 transition duration-150 shadow-md ${getButtonBg(
                      item.themeColor
                    )}`}
                  >
                    <span>Launch Utility</span>
                    <ExternalLink className="w-3.5 h-3.5" />
                  </a>
                </div>
              ))}
            </div>

            {/* Fast 1-Click Sample Data Bar */}
            <div className="rounded-2xl bg-gradient-to-r from-slate-800/90 via-[#001f3f]/90 to-slate-800/90 border border-slate-700 p-6 flex flex-col md:flex-row items-center justify-between gap-5 shadow-xl">
              <div className="flex items-center gap-4">
                <div className="w-12 h-12 rounded-xl bg-[#ffcc00]/20 border border-[#ffcc00]/40 flex items-center justify-center text-[#ffcc00] shrink-0">
                  <Zap className="w-6 h-6" />
                </div>
                <div>
                  <h4 className="text-sm font-bold text-white">⚡ Instant Abqaiq Sample Dataset</h4>
                  <p className="text-xs text-[#ffcc00] mt-0.5">
                    Pre-populate active session memory with official Abqaiq BI-10-10303 weekly observation records &amp; supervisor tables.
                  </p>
                </div>
              </div>

              <form action="/use-sample" method="post" className="shrink-0 w-full md:w-auto">
                <button
                  type="submit"
                  id="btnQuickLoadSampleData"
                  className="w-full md:w-auto flex items-center justify-center gap-2 bg-[#ffcc00] hover:bg-[#e6b800] active:scale-95 text-[#002244] font-bold text-xs px-5 py-3 rounded-xl shadow-lg transition duration-150"
                >
                  <RefreshCw className="w-4 h-4" />
                  <span>Load Sample Data &amp; Preview</span>
                </button>
              </form>
            </div>

          </main>

          {/* Footer */}
          <footer className="mt-auto border-t border-slate-800 bg-[#001428] px-6 py-4 text-center text-xs text-slate-400">
            <p>
              &copy; 2026 HEISCO – ENPPI – Saudi Aramco Abqaiq Facilities (BI-10-10303) &bull; Central HSE Operations Hub
            </p>
          </footer>
        </div>
      </div>
    </div>
  );
};

export default DashboardLayout;
