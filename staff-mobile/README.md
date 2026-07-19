# Adventure Centre Staff phone app

This Expo React Native app uses the same Supabase project and live tables as the manager app.

## Configure

```bash
cd staff-mobile
cp .env.example .env
```

Copy the values from the manager project's `.env.local`:

- `VITE_SUPABASE_URL` -> `EXPO_PUBLIC_SUPABASE_URL`
- `VITE_SUPABASE_ANON_KEY` -> `EXPO_PUBLIC_SUPABASE_ANON_KEY`

## Run on a phone with Expo Go

```bash
npm install
npx expo start
```

Scan the QR code with Expo Go.

## Build an installable Android APK

```bash
npx eas-cli@latest login
npx eas-cli@latest build:configure
npm run build:android
```

The preview profile in `eas.json` produces an APK. The production profile produces an Android App Bundle.
