# Music Label Manager Advanced v2

Advanced Flutter source project customized for your PDL distributor workflow.

## New features
- Persistent local data
- Client revenue shares
- Song/ISRC mapping
- PDL Excel import
- Client gross royalty calculation
- Label earnings calculation
- Recoverable advance tracking
- Advance recovery entries
- Payment and UTR/reference tracking
- Outstanding client balance
- Client account summary
- Dashboard

## Expected PDL columns
DSP Name
Month of Royalty
Song Name
ISRC
Income
Admin_Exp
Royalty Paid to PDL member

## Build APK
flutter create .
flutter pub get
flutter build apk --release

For production with very large reports and cloud backup, the next upgrade should move storage from SharedPreferences to SQLite/Drift or Supabase.
