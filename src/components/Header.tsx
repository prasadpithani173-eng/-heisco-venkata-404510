import React from 'react';

interface HeaderProps {
  className?: string;
}

export const Header: React.FC<HeaderProps> = ({ className = '' }) => {
  return (
    <header className={`w-full bg-white shadow-md border-b border-slate-200 overflow-hidden ${className}`} id="commonAppHeader">
      <div className="w-full mx-auto flex items-center justify-center p-0">
        <img
          src="/public/assets/header_final.png"
          alt="DEBOTTLENECK PRODUCTION FACILITIES - ABQAIQ | Enppi / HEISCO / Saudi Aramco"
          className="w-full max-h-24 md:max-h-28 lg:max-h-32 object-contain block mx-auto transition-all"
          onError={(e) => {
            const target = e.currentTarget as HTMLImageElement;
            if (!target.src.includes('/assets/header_final.png')) {
              target.src = '/assets/header_final.png';
            }
          }}
        />
      </div>
    </header>
  );
};

export default Header;
