import { Building2, CalendarDays, Home, NotebookPen, Settings, Users } from 'lucide-react'
import type { Page } from '../types'

const items = [
  { page: 'dashboard' as Page, label: 'Home', icon: Home },
  { page: 'programme' as Page, label: 'Programme', icon: CalendarDays },
  { page: 'staffing' as Page, label: 'Staffing', icon: Users },
  { page: 'arrivals' as Page, label: 'Arrivals', icon: Building2 },
  { page: 'schoolNotes' as Page, label: 'School Notes', icon: NotebookPen },
  { page: 'admin' as Page, label: 'Admin', icon: Settings },
]

export function Nav({ page, setPage }: { page: Page; setPage: (page: Page) => void }) {
  return (
    <nav className="top-nav">
      {items.map(({ page: itemPage, label, icon: Icon }) => {
        const active = page === itemPage || (itemPage === 'admin' && (page === 'staff' || page === 'holidays' || page === 'signoffs' || page === 'logs' || page === 'staffingLogs'))
        return (
          <button key={itemPage} className={active ? 'active' : ''} onClick={() => setPage(itemPage)}>
            <Icon size={19} />
            <span>{label}</span>
          </button>
        )
      })}
    </nav>
  )
}
