// src/components/user-management/user-management.component.ts  (UPDATED — pagination)
import { Component, ChangeDetectionStrategy, signal, computed, effect } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ApiService } from '../../services/api.service';
import { User, Role, AppScreen, TicketCategory, TICKET_CATEGORIES, USER_ROLES } from '../../models';
import { FormControl, FormGroup, ReactiveFormsModule, Validators } from '@angular/forms';
import { PaginationComponent } from '../shared/pagination.component';

/** Roles that can be assigned a category scope */
const CATEGORY_SCOPED_ROLES = [
  USER_ROLES.Admin,
  USER_ROLES.HardwareAdmin,
  USER_ROLES.SoftwareAdmin,
  USER_ROLES.SupportAgent,
];

@Component({
  selector: 'app-user-management',
  templateUrl: './user-management.component.html',
  styleUrls: ['./user-management.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [CommonModule, ReactiveFormsModule, PaginationComponent],
})
export class UserManagementComponent {
  users;
  roles;
  loading    = signal(false);
  apiLoading;

  activeTab       = signal<'users' | 'roles'>('users');
  isUserModalOpen = signal(false);
  isRoleModalOpen = signal(false);
  editingRole     = signal<Role | null>(null);
  editingUser     = signal<User | null>(null);

  availableScreens: AppScreen[] = [
    'Dashboard', 'Tickets', 'User Management', 'Reports', 'Dispatch', 'Customer Entry', 'HW Inventory', 'Price Configuration', 'Operations Tools',
  ];
  selectedPermissions = signal<boolean[]>(new Array(this.availableScreens.length).fill(false));

  ticketCategories: TicketCategory[] = TICKET_CATEGORIES;

  // ── Search / filter ─────────────────────────────────────────────────────────
  userSearch = signal('');

  filteredUsers = computed(() => {
    const q = this.userSearch().toLowerCase();
    if (!q) return this.users();
    return this.users().filter(u =>
      u.name.toLowerCase().includes(q) ||
      u.username.toLowerCase().includes(q)
    );
  });

  // ── Pagination — Users tab ──────────────────────────────────────────────────
  usersPage = signal(1);
  readonly usersPageSize = 15;

  pagedUsers = computed(() => {
    const start = (this.usersPage() - 1) * this.usersPageSize;
    return this.filteredUsers().slice(start, start + this.usersPageSize);
  });

  // ── Pagination — Roles tab ──────────────────────────────────────────────────
  rolesPage = signal(1);
  readonly rolesPageSize = 15;

  pagedRoles = computed(() => {
    const start = (this.rolesPage() - 1) * this.rolesPageSize;
    return this.roles().slice(start, start + this.rolesPageSize);
  });

  // Reset users page when search changes
  private resetUsersPage = effect(() => {
    this.userSearch();
    this.usersPage.set(1);
  }, { allowSignalWrites: true });

  // ── Forms ───────────────────────────────────────────────────────────────────
  newUserForm = new FormGroup({
    name:     new FormControl('', Validators.required),
    username: new FormControl('', Validators.required),
    password: new FormControl('', [Validators.required, Validators.minLength(8)]),
    roleId:   new FormControl<number | null>(null, Validators.required),
    category: new FormControl<TicketCategory | null>(null),
  });

  roleForm = new FormGroup({
    name: new FormControl('', Validators.required),
  });

  editUserForm = new FormGroup({
    roleId:   new FormControl<number | null>(null, Validators.required),
    category: new FormControl<TicketCategory | null>(null),
  });

  newUserNeedsCategory = computed(() => {
    const roleId   = this.newUserForm.value.roleId;
    const roleName = this.roles().find(r => r.id === roleId)?.name ?? '';
    return CATEGORY_SCOPED_ROLES.includes(roleName as any);
  });

  editUserNeedsCategory = computed(() => {
    const roleId   = this.editUserForm.value.roleId;
    const roleName = this.roles().find(r => r.id === roleId)?.name ?? '';
    return CATEGORY_SCOPED_ROLES.includes(roleName as any);
  });

  constructor(private apiService: ApiService) {
    this.users      = this.apiService.users;
    this.roles      = this.apiService.roles;
    this.apiLoading = this.apiService.loading;
  }

  getRoleById(id: number)  { return this.roles().find(r => r.id === id); }

  getCategoryBadgeStyle(c?: TicketCategory | null) {
    if (!c) return '';
    const map: Record<TicketCategory, string> = {
      Hardware: 'background:#EFF6FF; color:#1d4ed8;',
      Software: 'background:#F0FDF4; color:#15803d;',
      ASRS:     'background:#FFF7ED; color:#c2410c;',
    };
    return map[c] ?? '';
  }

  selectTab(tab: 'users' | 'roles' | EventTarget | null) {
    const v = (typeof tab === 'string') ? tab : (tab as HTMLSelectElement)?.value;
    if (v === 'users' || v === 'roles') this.activeTab.set(v);
  }

  openUserModal() {
    this.newUserForm.reset({ roleId: this.roles()[0]?.id ?? null, category: null });
    this.isUserModalOpen.set(true);
  }

  closeUserModal() { this.isUserModalOpen.set(false); this.editingUser.set(null); }

  async createUser() {
    if (this.newUserForm.invalid) return;
    this.loading.set(true);
    try {
      const v = this.newUserForm.getRawValue();
      await this.apiService.createUser({
        name:         v.name!,
        username:     v.username!,
        contactEmail: '',
        roleId:       v.roleId!,
        category:     this.newUserNeedsCategory() ? v.category : null,
      } as any);
      this.closeUserModal();
    } catch (err) {
      console.error('Failed to create user', err);
    } finally {
      this.loading.set(false);
    }
  }

  openEditUserModal(user: User) {
    this.editingUser.set(user);
    this.editUserForm.patchValue({ roleId: user.roleId, category: user.category ?? null });
    this.isUserModalOpen.set(true);
  }

  async saveEditUser() {
    if (this.editUserForm.invalid) return;
    const user = this.editingUser();
    if (!user) return;
    this.loading.set(true);
    try {
      const v = this.editUserForm.getRawValue();
      await this.apiService.updateUser({
        ...user,
        roleId:   v.roleId!,
        category: this.editUserNeedsCategory() ? v.category : null,
      });
      this.closeUserModal();
    } catch (err) {
      console.error('Failed to update user', err);
    } finally {
      this.loading.set(false);
    }
  }

  openRoleModal(role?: Role) {
    this.editingRole.set(role ?? null);
    if (role) {
      this.roleForm.patchValue({ name: role.name });
      const perms = new Array(this.availableScreens.length).fill(false);
      role.permissions.forEach(p => {
        const i = this.availableScreens.indexOf(p);
        if (i >= 0) perms[i] = true;
      });
      this.selectedPermissions.set(perms);
    } else {
      this.roleForm.reset();
      this.selectedPermissions.set(new Array(this.availableScreens.length).fill(false));
    }
    this.isRoleModalOpen.set(true);
  }

  closeRoleModal() { this.isRoleModalOpen.set(false); this.editingRole.set(null); }

  togglePermission(index: number, checked: boolean) {
    const perms = [...this.selectedPermissions()];
    perms[index] = checked;
    this.selectedPermissions.set(perms);
  }

  async saveRole() {
    if (this.roleForm.invalid) return;
    this.loading.set(true);
    try {
      const permissions = this.availableScreens.filter((_, i) => this.selectedPermissions()[i]);
      const role = this.editingRole();
      if (role) {
        await this.apiService.updateRole({ ...role, name: this.roleForm.value.name!, permissions });
      } else {
        await this.apiService.createRole({ name: this.roleForm.value.name!, permissions } as any);
      }
      this.closeRoleModal();
    } catch (err) {
      console.error('Failed to save role', err);
    } finally {
      this.loading.set(false);
    }
  }
}
