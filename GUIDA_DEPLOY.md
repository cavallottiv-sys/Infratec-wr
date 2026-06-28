# 🚀 GUIDA DEPLOY — INFRATEC WR Manager
## Da zero a link pubblico in ~20 minuti

---

## STEP 1 — Crea account GitHub (se non ce l'hai)
1. Vai su https://github.com
2. Clicca **Sign up** → email, password, username
3. Verifica l'email

---

## STEP 2 — Crea repository GitHub
1. Accedi a GitHub → clicca **+** in alto a destra → **New repository**
2. **Repository name:** `infratec-wr` (esatto, tutto minuscolo)
3. Spunta **Public**
4. Clicca **Create repository**
5. **Copia il tuo username GitHub** — ti serve dopo

---

## STEP 3 — Crea progetto Firebase
1. Vai su https://console.firebase.google.com (con il tuo account Google)
2. Clicca **Aggiungi progetto** → nome: `infratec-wr`
3. Disabilita Google Analytics (non serve) → **Crea progetto**
4. Nel menu a sinistra → **Realtime Database** → **Crea database**
   - Scegli regione: **europe-west1**
   - Modalità: **Test** → OK
5. Vai in ⚙ **Impostazioni progetto** (ingranaggio in alto a sinistra)
6. Scorri fino a **Le tue app** → clicca icona **</>** (Web)
7. Nome app: `infratec-wr` → **Registra app**
8. **COPIA** il blocco `firebaseConfig` che compare (lo usi nel passo 5)

---

## STEP 4 — Installa Node.js (se non ce l'hai)
- Vai su https://nodejs.org → scarica la versione **LTS**
- Installa normalmente

---

## STEP 5 — Configura e build del progetto

### 5a. Estrai lo ZIP scaricato da Claude in una cartella, es. `C:\infratec-wr`

### 5b. Apri il file `src/firebase.js` con Blocco Note e incolla i tuoi dati:
```js
const firebaseConfig = {
  apiKey:            "LA TUA API KEY",
  authDomain:        "infratec-wr.firebaseapp.com",
  databaseURL:       "https://infratec-wr-default-rtdb.europe-west1.firebasedatabase.app",
  projectId:         "infratec-wr",
  storageBucket:     "infratec-wr.appspot.com",
  messagingSenderId: "123456789",
  appId:             "1:123456789:web:abc123",
};
```

### 5c. Apri `vite.config.js` e verifica che ci sia il tuo nome repo:
```js
base: '/infratec-wr/',   // deve corrispondere al nome repository GitHub
```

### 5d. Apri il **Prompt dei comandi** (cmd) nella cartella del progetto:
```
cd C:\infratec-wr
npm install
npm run build
```
Aspetta che finisca (1-2 minuti). Verrà creata la cartella `dist/`.

---

## STEP 6 — Pubblica su GitHub Pages

Nel prompt dei comandi, nella stessa cartella:

```bash
git init
git add .
git commit -m "primo deploy"
git branch -M main
git remote add origin https://github.com/TUO-USERNAME/infratec-wr.git
git push -u origin main
```
*(sostituisci TUO-USERNAME con il tuo username GitHub)*

Poi pubblica la cartella `dist/`:
```bash
npx gh-pages -d dist
```

---

## STEP 7 — Abilita GitHub Pages
1. Vai su GitHub → repo `infratec-wr` → **Settings**
2. Menu sinistra → **Pages**
3. **Source:** Deploy from a branch
4. **Branch:** `gh-pages` → `/root` → **Save**

Dopo 1-2 minuti il link sarà:
```
https://TUO-USERNAME.github.io/infratec-wr/
```

---

## STEP 8 — Aggiornamenti futuri
Ogni volta che modifichi il codice:
```bash
npm run deploy
```
Un solo comando — fa build + pubblica automaticamente.

---

## ✅ Checklist finale
- [ ] Node.js installato
- [ ] Account GitHub creato + repo `infratec-wr`
- [ ] Progetto Firebase creato + Realtime Database in europe-west1
- [ ] `src/firebase.js` compilato con i tuoi dati
- [ ] `npm install` eseguito senza errori
- [ ] `npm run build` eseguito senza errori
- [ ] `npx gh-pages -d dist` eseguito
- [ ] GitHub Pages abilitato su branch `gh-pages`
- [ ] Link funzionante su tutti i dispositivi 🎉

---

## ❓ Problemi comuni

**"npm non riconosciuto"** → Node.js non è installato o riavvia il cmd dopo l'installazione

**"permission denied" su git push** → GitHub chiede autenticazione. Vai su GitHub → Settings → Developer Settings → Personal Access Tokens → Genera token con permesso `repo` e usalo come password

**App apre ma i dati non si salvano** → Controlla `src/firebase.js`: il campo `databaseURL` deve essere presente e corretto

**Pagina bianca** → `vite.config.js`: controlla che `base` corrisponda ESATTAMENTE al nome del repository GitHub
