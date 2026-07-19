import { Building2, CalendarDays, Home, NotebookPen, Settings, Users } from 'lucide-react'
import type { AccountRole, Page } from '../types'

const items = [
  { page: 'dashboard' as Page, label: 'Home', icon: Home },
  { page: 'programme' as Page, label: 'Programme', icon: CalendarDays },
  { page: 'staffing' as Page, label: 'Staffing', icon: Users },
  { page: 'arrivals' as Page, label: 'Arrivals', icon: Building2 },
  { page: 'schoolNotes' as Page, label: 'School Notes', icon: NotebookPen },
  { page: 'admin' as Page, label: 'Admin', icon: Settings },
]

export function Nav({ page, setPage, accountRole }: { page: Page; setPage: (page: Page) => void; accountRole: AccountRole }) {
  const visibleItems = accountRole === 'admin'
    ? items.filter((item) => item.page === 'admin')
    : items

  return (
    <nav className="top-nav">
      {visibleItems.map(({ page: itemPage, label, icon: Icon }) => {
        const active = page === itemPage || (itemPage === 'admin' && (page === 'staff' || page === 'holidays' || page === 'signoffs' || page === 'logs' || page === 'staffingLogs' || page === 'programmeArchive'))
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
