# IT Onboarding Form - Specification

## 1. Project Overview
- **Project Name**: IT Onboarding Portal
- **Type**: React.js Web Application with MSAL Authentication
- **Core Functionality**: A multi-step onboarding form accessible only after Microsoft 365 login, with dark/light mode and shareable QR code/link functionality
- **Target Users**: New employees requiring IT onboarding

## 2. UI/UX Specification

### Layout Structure
- **Login Page**: Centered login card with Microsoft login button
- **Homepage**: Empty initial landing page (redirects to form)
- **Form Page**: Full-width layout with:
  - Fixed top auth banner showing user info
  - Main content area (centered, max-width 800px)
  - Floating action buttons for QR/share

### Responsive Breakpoints
- Mobile: < 768px
- Tablet: 768px - 1024px
- Desktop: > 1024px

### Visual Design
- **Color Palette - Light Mode**:
  - Background: `#FAFAFA` (off-white)
  - Surface: `#FFFFFF` (pure white)
  - Primary Text: `#1A1A1A` (near-black)
  - Secondary Text: `#6B6B6B` (gray)
  - Accent/Primary: `#000000` (black)
  - Accent Hover: `#333333`
  - Border: `#E5E5E5`
  - Success: `#10B981`
  - Error: `#EF4444`

- **Color Palette - Dark Mode**:
  - Background: `#0A0A0A` (near-black)
  - Surface: `#141414` (dark gray)
  - Primary Text: `#FAFAFA` (off-white)
  - Secondary Text: `#A3A3A3` (light gray)
  - Accent/Primary: `#FFFFFF` (white)
  - Accent Hover: `#E5E5E5`
  - Border: `#262626`

- **Typography**:
  - Font Family: `'Inter', -apple-system, BlinkMacSystemFont, sans-serif`
  - Headings: 600 weight
    - H1: 32px
    - H2: 24px
    - H3: 18px
  - Body: 400 weight, 16px
  - Small: 14px

- **Spacing System**: 4px base unit (4, 8, 12, 16, 24, 32, 48, 64)

- **Visual Effects**:
  - Border radius: 12px (cards), 8px (buttons), 6px (inputs)
  - Shadows (light mode): `0 1px 3px rgba(0,0,0,0.08), 0 4px 12px rgba(0,0,0,0.04)`
  - Shadows (dark mode): `0 1px 3px rgba(0,0,0,0.3)`
  - Transitions: 200ms ease-out

### Components

1. **Auth Banner** (Form Page Top)
   - User avatar circle (initials)
   - User email/display name
   - Sign out button
   - Theme toggle (sun/moon icon)

2. **Login Card**
   - App logo/title
   - "Sign in with Microsoft" button (styled black/white)
   - Subtle description text

3. **Onboarding Form (SurveyJS)**
   - Multi-step form wizard
   - Progress indicator
   - Fields: Personal Info, Equipment Request, Access Permissions
   - Navigation buttons (Previous/Next/Submit)

4. **Share Panel** (Floating)
   - QR Code display (generated via library)
   - "Copy Link" button
   - "Download QR" button
   - Toggle button (floating action button)

## 3. Functionality Specification

### Authentication Flow
1. User visits homepage → redirects to Login page
2. Login page shows MSAL login button
3. User clicks "Sign in with Microsoft"
4. On success → redirect to Form page
5. Form page shows auth banner with user info
6. Session persists page-to-page via MSAL cache

### Form Page Features
- **Theme Toggle**: Dark/Light mode switch (persists in localStorage)
- **SurveyJS Form**:
  - Step 1: Personal Information (name, department, role, start date)
  - Step 2: Equipment Needs (laptop, monitors, peripherals checkboxes)
  - Step 3: Access Requests (email, Teams, SharePoint, VPN checkboxes)
  - Step 4: Review & Submit
- **Share Functionality**:
  - Generate QR code containing current URL
  - Copy URL to clipboard
  - Download QR as PNG

### Edge Cases
- Handle MSAL login failure gracefully
- Form validation before step progression
- Preserve form data in session storage

## 4. Acceptance Criteria
- [ ] Login page displays with Microsoft sign-in button
- [ ] Successful Microsoft login redirects to Form page
- [ ] Form page shows user info in top banner
- [ ] Dark/Light mode toggle works and persists
- [ ] SurveyJS form renders with all 4 steps
- [ ] QR code generates and displays correctly
- [ ] Copy link copies URL to clipboard
- [ ] Download QR saves as image file
- [ ] Responsive design works on mobile/tablet/desktop
- [ ] Sleek white/black aesthetic achieved