export type UserRole = 'admin' | 'officer' | 'supervisor' | 'engineer';

export interface UserSession {
  username: string;
  role: UserRole;
  projectCode: string;
  facility: string;
}

export interface HSEUtilityLink {
  id: string;
  title: string;
  description: string;
  href: string;
  icon: string; // Lucide icon name or emoji
  badge: string;
  format: 'Word (.docx)' | 'Excel (.xlsx)' | 'PowerPoint (.pptx)' | 'Analytics' | 'System';
  tags: string[];
  themeColor: 'green' | 'teal' | 'navy' | 'amber';
  adminOnly?: boolean;
}

export interface NavItem {
  id: string;
  label: string;
  href: string;
  icon: string;
  badge?: string;
  adminOnly?: boolean;
}
