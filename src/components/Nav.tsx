import { Building2, CalendarDays, CheckSquare, Home, Users } from 'lucide-react'
import type { Page } from '../types'

const items = [
  { page: 'dashboard' as Page, label: 'Home', icon: Home },
  { page: 'programme' as Page, label: 'Programme', icon: CalendarDays },
  { page: 'staffing' as Page, label: 'Staffing', icon: Users },
  { page: 'accommodation' as Page, label: 'Accommodation', icon: Building2 },
  { page: 'staff' as Page, label: 'Staff', icon: Users },
  { page: 'signoffs' as Page, label: 'Sign-offs', icon: CheckSquare },
]

export function Nav({
  page,
  setPage,
}: {
  page: Page
  setPage: (page: Page) => void
}) {
  return (
    <nav className="top-nav">
      {items.map(({ page: itemPage, label, icon: Icon }) => (
        <button
          key={itemPage}
          className={page === itemPage ? 'active' : ''}
          onClick={() => setPage(itemPage)}
        >
          <Icon size={19} />
          <span>{label}</span>
        </button>
      ))}
    </nav>
  )
}
