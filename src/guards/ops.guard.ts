// src/guards/ops.guard.ts
//
// Gates the whole /ops/* subtree (the auto-discovered "Operations Tools"
// feature pages — ironing/packing, shipment dashboards, box packing, etc.).
// Access is granted to Admins, Managers, or anyone granted the explicit
// 'Operations Tools' AppScreen permission — matching canViewOperations in
// layout.component.ts.
import { Injectable } from '@angular/core';
import { CanActivate, Router, UrlTree } from '@angular/router';
import { AuthService } from '../services/auth.service';

@Injectable({ providedIn: 'root' })
export class OpsGuard implements CanActivate {
  constructor(
    private authService: AuthService,
    private router: Router,
  ) {}

  canActivate(): boolean | UrlTree {
    if (
      this.authService.hasPermission('Operations Tools') ||
      this.authService.isAdmin() ||
      this.authService.isManager()
    ) {
      return true;
    }
    return this.router.parseUrl('/dashboard');
  }
}
