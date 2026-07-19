import { Building2, CalendarDays, Home, NotebookPen, Settings, Users } from 'lucide-react'
import type { Page } from '../types'

const items = [
  { page: 'dashboard' as Page, label: 'Dashboard', icon: Home },
  { page: 'programme' as Page, label: 'Programme', icon: CalendarDays },
  { page: 'staffing' as Page, label: 'Staffing', icon: Users },
  { page: 'arrivals' as Page, label: 'Arrivals', icon: Building2 },
  { page: 'schoolNotes' as Page, label: 'School Notes', icon: NotebookPen },
  { page: 'admin' as Page, label: 'Admin', icon: Settings },
]

export function Nav({ page, setPage }: { page: Page; setPage: (page: Page) => void }) {
  return (
    <aside className="app-sidebar">
      <div className="sidebar-brand">
        <img src={`${import.meta.env.BASE_URL}manor-adventure-logo.png`} alt="Manor Adventure" />
        <div>
          <span>Norfolk Lakes</span>
          <strong>Centre Manager</strong>
        </div>
      </div>
      <p className="sidebar-section-label">Workspace</p>
      <nav className="top-nav" aria-label="Main navigation">
        {items.map(({ page: itemPage, label, icon: Icon }) => {
          const active = page === itemPage || (itemPage === 'admin' && ['staff', 'holidays', 'signoffs', 'logs', 'staffingLogs', 'programmeBuilder', 'activitiesEquipment', 'programmeArchive', 'formerStaff', 'loanHistory'].includes(page))
          return (
            <button key={itemPage} className={active ? 'active' : ''} onClick={() => setPage(itemPage)}>
              <Icon size={20} />
              <span>{label}</span>
            </button>
          )
        })}
      </nav>
      <div className="sidebar-footer">
        <span className="sidebar-status-dot" />
        <div><strong>System online</strong><small>Version 3.0</small></div>
      </div>
    </aside>
  )
}
