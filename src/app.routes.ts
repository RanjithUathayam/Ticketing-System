// src/app.routes.ts  (UPDATED — customer-entry routes added)
import { Routes } from '@angular/router';
import { AuthGuard } from './guards/auth.guard';
import { OpsGuard } from './guards/ops.guard';
import { featureRoutes } from './features/features.generated';

export const appRoutes: Routes = [
  {
    path: 'login',
    loadComponent: () => import('./components/login/login.component').then(c => c.LoginComponent)
  },
  {
    path: '',
    canActivate: [AuthGuard],
    canActivateChild: [AuthGuard],
    loadComponent: () => import('./components/layout/layout.component').then(c => c.LayoutComponent),
    children: [
      { path: '', redirectTo: 'dashboard', pathMatch: 'full' },
      {
        path: 'dashboard',
        loadComponent: () => import('./components/dashboard/dashboard.component').then(c => c.DashboardComponent)
      },
      {
        path: 'tickets',
        loadComponent: () => import('./components/ticket-list/ticket-list.component').then(c => c.TicketListComponent)
      },
      {
        path: 'tickets/:id',
        loadComponent: () => import('./components/ticket-detail/ticket-detail.component').then(c => c.TicketDetailComponent)
      },
      {
        path: 'user-management',
        loadComponent: () => import('./components/user-management/user-management.component').then(c => c.UserManagementComponent)
      },
      {
        path: 'reports',
        loadComponent: () => import('./components/reports/reports.component').then(c => c.ReportsComponent)
      },
      // ── Online Dispatch ──────────────────────────────────────────────────
      {
        path: 'dispatch',
        loadComponent: () => import('./components/dispatch-list/dispatch-list.component').then(c => c.DispatchListComponent)
      },
      {
        path: 'dispatch/:id',
        loadComponent: () => import('./components/dispatch-form/dispatch-form.component').then(c => c.DispatchFormComponent)
      },
      // ── Online Customer Entry ────────────────────────────────────────────
      {
        path: 'customer-entry',
        loadComponent: () => import('./components/customer-entry-list/customer-entry-list.component').then(c => c.CustomerEntryListComponent)
      },
      {
        path: 'customer-entry/:id',
        loadComponent: () => import('./components/customer-entry-form/customer-entry-form.component').then(c => c.CustomerEntryFormComponent)
      },
      {
        path: 'price-configuration',
        loadComponent: () => import('./components/price-configuration/price-configuration.component').then(c => c.PriceConfigurationComponent)
      },
      {
        path: 'label-print-config',
        loadComponent: () => import('./components/label-print-config/label-print-config.component').then(c => c.LabelPrintConfigComponent)
      },
      {
        path: 'label-print',
        loadComponent: () => import('./components/label-print/label-print.component').then(c => c.LabelPrintComponent)
      },
      {
        path: 'hw-inventory',
        loadComponent: () => import('./components/hw-inventory-list/hw-inventory-list.component').then(c => c.HwInventoryListComponent)
        },
        {
            path: 'hw-inventory/:id',
            loadComponent: () => import('./components/inventory-form/inventory-form.component').then(c => c.HwInventoryFormComponent)
        },
      // ── Operations Tools (auto-discovered — see src/features/README.md) ──
      // New pages under src/features/<key>/ appear here automatically after
      // `npm run generate:features` (wired into predev/prebuild); this
      // single block is the only manual wiring auto-discovery ever needs.
      {
        path: 'ops',
        canActivate: [OpsGuard],
        children: featureRoutes,
      },
    ]
  },
  { path: '**', redirectTo: '/login' }
];
