# Admin Portal Implementation Plan

## Goal
Enable direct backend login via the Web App interface, allowing instructors to manage the system without navigating Google Drive.

## Proposed Changes

### 1. Security (`Service_Security.gs`)
- Add `verifyAdmin(password)` function.
- Storage: Use `ScriptProperties` for 'ADMIN_PASSWORD' (default: 'admin123').

### 2. Backend (`Main.gs`)
- **Route**: Handle `?route=admin`.
- **API**:
    - `apiAdminLogin(password)`: Returns token/success.
    - `apiGetAdminDashboardData()`: Returns active activity stats, links to sheets.

### 3. Frontend (`UI_Admin.html`)
- **Login View**: Simple password field.
- **Dashboard View**:
    - Show "Current Activity".
    - Button: "Open Control Sheet" (Link).
    - Button: "Open Data Sheet" (Link).
    - Stats: "Total Submissions", "Completion Rate".

### 4. Integration
- Add `UI_Admin` to `Main.gs` template loader.
