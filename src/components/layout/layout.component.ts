// src/components/layout/layout.component.ts
import { Component, ChangeDetectionStrategy, signal, computed } from '@angular/core';
import { RouterOutlet, RouterLink, RouterLinkActive, Router, NavigationEnd } from '@angular/router';
import { CommonModule } from '@angular/common';
import { filter } from 'rxjs';
import { AuthService } from '../../services/auth.service';
import { NotificationService } from '../../services/notification.service';
import { FEATURE_MENU } from '../../features/features.generated';
import { FEATURE_ICON_PATHS } from '../../features/feature-icons';
import { FeaturePageMeta } from '../../features/feature.types';

@Component({
  selector: 'app-layout',
  templateUrl: './layout.component.html',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [CommonModule, RouterOutlet, RouterLink, RouterLinkActive],
})
export class LayoutComponent {
  user;
  userRole;
  notifications;
  unreadCount;

  userMenuOpen             = signal(false);
  notificationPanelOpen    = signal(false);
  sidebarOpen              = signal(false);
  desktopSidebarCollapsed  = signal(false);
  pageTitle                = signal('Dashboard');

  sessionReady;

  canViewDashboard;
  canViewTickets;
  canViewUserManagement;
  canViewReports;
  canViewDispatch;
  canViewCustomerEntry;
  canViewPriceConfiguration;
  canViewLabelPrintConfig;
  canViewLabelPrint;
  canViewHwInventory;
  canViewOperations;

  /** Auto-discovered "Operations Tools" pages, grouped by their `group` field and sorted by `order`. See src/features/README.md. */
  readonly opsGroups: { group: string; items: FeaturePageMeta[] }[];
  readonly featureIconPath = (icon: FeaturePageMeta['icon']) => FEATURE_ICON_PATHS[icon] ?? FEATURE_ICON_PATHS.default;

  constructor(
    public authService:         AuthService,
    public notificationService: NotificationService,
    public router:              Router,
  ) {
    this.user          = this.authService.currentUser;
    this.userRole      = this.authService.currentUserRole;
    this.notifications = this.notificationService.notifications;
    this.unreadCount   = this.notificationService.unreadCount;

    this.sessionReady = this.authService.sessionReady;

    this.canViewDashboard      = computed(() => this.authService.hasPermission('Dashboard'));
    this.canViewTickets        = computed(() => this.authService.hasPermission('Tickets'));
    // User Management: visible to Admin, Manager (same access level as Operations
    // Tools — see canViewOperations below), or anyone with explicit permission.
    this.canViewUserManagement = computed(() =>
      this.authService.hasPermission('User Management') ||
      this.authService.isAdmin() ||
      this.authService.isManager()
    );
    this.canViewReports        = computed(() => this.authService.hasPermission('Reports'));
    this.canViewDispatch       = computed(() =>
      this.authService.hasPermission('Dispatch') ||
      this.authService.isAdmin() ||
      this.authService.isManager() ||
      this.authService.isSupport()
    );
    this.canViewCustomerEntry  = computed(() =>
      this.authService.hasPermission('Customer Entry') ||
      this.authService.isAdmin() ||
      this.authService.isManager()
    );
    this.canViewPriceConfiguration = computed(() =>
      this.authService.hasPermission('Price Configuration') ||
      this.authService.isAdmin() ||
      this.authService.isManager()
    );
    this.canViewLabelPrintConfig = computed(() =>
      this.authService.hasPermission('Price Configuration') ||
      this.authService.isAdmin() ||
      this.authService.isManager()
    );
    this.canViewLabelPrint = computed(() =>
      this.authService.hasPermission('Price Configuration') ||
      this.authService.isAdmin() ||
      this.authService.isManager()
    );
    // HW Inventory: visible to Admin, Manager, or anyone with explicit permission
    this.canViewHwInventory    = computed(() =>
      this.authService.hasPermission('HW Inventory') ||
      this.authService.isAdmin() ||
      this.authService.isManager()
    );
    // Operations Tools (auto-discovered pages under src/features/) — Admin/Manager,
    // or anyone with explicit permission. Matches OpsGuard in src/guards/ops.guard.ts.
    this.canViewOperations = computed(() =>
      this.authService.hasPermission('Operations Tools') ||
      this.authService.isAdmin() ||
      this.authService.isManager()
    );

    this.opsGroups = Object.values(
      FEATURE_MENU
        .filter(m => m.enabled)
        .reduce((acc, m) => {
          (acc[m.group] ||= { group: m.group, items: [] }).items.push(m);
          return acc;
        }, {} as Record<string, { group: string; items: FeaturePageMeta[] }>)
    );

    this.router.events.pipe(
      filter((e): e is NavigationEnd => e instanceof NavigationEnd)
    ).subscribe((e: NavigationEnd) => {
      this.pageTitle.set(this.getRouteTitle(e.urlAfterRedirects));
      this.sidebarOpen.set(false);
    });
  }

  getRouteTitle(url: string): string {
    if (url.includes('/dashboard'))           return 'Dashboard';
    if (url.includes('/tickets/'))            return 'Ticket Details';
    if (url.includes('/tickets'))             return 'Tickets';
    if (url.includes('/user-management'))     return 'User Management';
    if (url.includes('/reports'))             return 'Reports';
    if (url.match(/\/dispatch\/\d+/))         return 'Dispatch Entry';
    if (url.includes('/dispatch/new'))        return 'New Dispatch Entry';
    if (url.includes('/dispatch'))            return 'Online Dispatch Details';
    if (url.match(/\/customer-entry\/\d+/))   return 'Customer Entry';
    if (url.includes('/customer-entry/new'))  return 'New Customer Entry';
    if (url.includes('/customer-entry'))      return 'Online Customer Entry';
    if (url.includes('/price-configuration')) return 'Price Configuration';
    if (url.includes('/label-print-config'))  return 'Label Print Configuration';
    if (url.match(/\/hw-inventory\/\d+/))     return 'Edit Asset';
    if (url.includes('/hw-inventory/new'))    return 'New Asset';
    if (url.includes('/hw-inventory'))        return 'HW Inventory';
    if (url.includes('/ops/')) {
      const segment = url.split('/ops/')[1]?.split(/[/?#]/)[0];
      const feature = FEATURE_MENU.find(m => m.path === segment);
      if (feature) return feature.title;
    }
    return 'Dashboard';
  }

  logout() {
    this.authService.logout();
    this.router.navigate(['/login']);
  }

  toggleSidebar()              { this.sidebarOpen.update(v => !v); }
  toggleDesktopSidebar()       { this.desktopSidebarCollapsed.update(v => !v); }
  toggleUserMenu()             { this.userMenuOpen.update(v => !v); }
  toggleNotificationPanel()    { this.notificationPanelOpen.update(v => !v); }
}
