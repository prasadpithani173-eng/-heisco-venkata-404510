import React from 'react';

export default function Header({ className = '' }) {
  return (
    <header className={`w-full bg-white shadow-md border-b border-slate-200 overflow-hidden ${className}`} id="commonAppHeader">
      <div className="w-full mx-auto flex items-center justify-center p-0">
        <img
          src="/public/assets/header_final.png"
          alt="DEBOTTLENECK PRODUCTION FACILITIES - ABQAIQ | Enppi / HEISCO / Saudi Aramco"
          className="w-full max-h-24 md:max-h-28 lg:max-h-32 object-contain block mx-auto transition-all"
          onError={(e) => {
            // Fallback path if loaded via vite or root static
            if (!e.currentTarget.src.includes('/assets/header_final.png')) {
              e.currentTarget.src = '/assets/header_final.png';
            }
          }}
        />
      </div>
    </header>
  );
}
