// src/guards/ops.guard.ts
//
// Gates the whole /ops/* subtree (the auto-discovered "Operations Tools"
// feature pages — ironing/packing, shipment dashboards, box packing, etc.).
// These pages aren't part of the existing per-screen AppScreen permission
// enum (that would require a backend/DB migration), so — matching the same
// fallback already used for Dispatch, Customer Entry, Price Configuration
// and HW Inventory in layout.component.ts — access is limited to Admins and
// Managers.
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
    if (this.authService.isAdmin() || this.authService.isManager()) {
      return true;
    }
    return this.router.parseUrl('/dashboard');
  }
}
