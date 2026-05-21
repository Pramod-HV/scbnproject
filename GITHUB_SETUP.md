# GitHub Authentication Setup

## 🔐 Fixing "Authentication failed" Error

You got this error:
```
remote: Invalid username or token. Password authentication is not supported for Git operations.
fatal: Authentication failed for 'https://github.com/Pramod-HV/scbnproject.git/'
```

This is because GitHub requires a **Personal Access Token (PAT)** instead of your password.

---

## ✅ Solution: Create a Personal Access Token

### Step 1: Create Personal Access Token on GitHub

1. **Go to GitHub.com** and log in
2. **Click your profile picture** (top right) → **Settings**
3. **Scroll down** to **Developer settings** (bottom left)
4. **Click "Personal access tokens"** → **"Tokens (classic)"**
5. **Click "Generate new token"** → **"Generate new token (classic)"**
6. **Fill in the form**:
   - **Note**: "Sterling Map Migrator - Mac"
   - **Expiration**: 90 days (or "No expiration" for testing)
   - **Select scopes**: Check these boxes:
     - ✅ `repo` (Full control of private repositories)
     - ✅ `workflow` (Update GitHub Action workflows)
7. **Click "Generate token"** (green button at bottom)
8. **IMPORTANT**: Copy the token immediately! It looks like:
   ```
   ghp_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
   ```
   **You won't be able to see it again!**

---

## 🔑 Step 2: Use Token to Push

Now use the token instead of your password:

```bash
# Push to GitHub (it will ask for credentials)
git push -u origin main

# When prompted:
# Username: Pramod-HV
# Password: [paste your token here - ghp_xxxxxxxxxxxx]
```

**OR** use this command to include credentials in URL (easier):

```bash
# Replace YOUR_TOKEN with the token you copied
git remote set-url origin https://YOUR_TOKEN@github.com/Pramod-HV/scbnproject.git

# Now push
git push -u origin main
```

---

## 🎯 Complete Steps from Where You Left Off

```bash
# You're currently here:
cd /Users/pramodhv/Documents/projectSCBN/Ai-Powered-Map-Migration-Accelarators/Integrated_Folder

# Option 1: Push with token in URL (easiest)
git remote set-url origin https://YOUR_TOKEN@github.com/Pramod-HV/scbnproject.git
git push -u origin main

# Option 2: Push and enter token when prompted
git push -u origin main
# Username: Pramod-HV
# Password: [paste token]
```

---

## 🔄 Alternative: Use SSH Instead of HTTPS

If you prefer SSH (more secure, no token needed each time):

### Step 1: Generate SSH Key

```bash
# Generate SSH key
ssh-keygen -t ed25519 -C "your_email@example.com"

# Press Enter for default location
# Press Enter for no passphrase (or set one)

# Copy the public key
cat ~/.ssh/id_ed25519.pub
```

### Step 2: Add SSH Key to GitHub

1. Go to GitHub.com → Settings
2. Click "SSH and GPG keys"
3. Click "New SSH key"
4. Paste the key you copied
5. Click "Add SSH key"

### Step 3: Change Remote URL to SSH

```bash
# Change from HTTPS to SSH
git remote set-url origin git@github.com:Pramod-HV/scbnproject.git

# Push
git push -u origin main
```

---

## 📋 Quick Reference

### Current Status:
- ✅ Git initialized
- ✅ Files committed
- ✅ Remote added
- ❌ Push failed (authentication)

### What You Need:
1. Personal Access Token from GitHub
2. Use token as password when pushing

### Commands:
```bash
# Get your token from GitHub (see Step 1 above)

# Then either:
# Method A: Include token in URL
git remote set-url origin https://YOUR_TOKEN@github.com/Pramod-HV/scbnproject.git
git push -u origin main

# Method B: Enter token when prompted
git push -u origin main
# Username: Pramod-HV
# Password: [paste token]
```

---

## ✅ After Successful Push

Once you successfully push:

1. **Go to your GitHub repo**: https://github.com/Pramod-HV/scbnproject
2. **Click "Actions" tab**
3. **You'll see the workflow running** (orange dot)
4. **Wait 5-10 minutes** for it to complete (green checkmark)
5. **Click on the workflow run**
6. **Download the installer** from "Artifacts" section
7. **Extract the ZIP** to get `SterlingMapMigrator_Setup_v1.0.0.exe`
8. **Distribute to Windows users!**

---

## 🆘 Troubleshooting

### "Repository not found"
- Make sure the repo exists on GitHub
- Check the URL is correct
- Verify you have access to the repo

### "Permission denied"
- Token doesn't have correct permissions
- Regenerate token with `repo` and `workflow` scopes

### "Token expired"
- Generate a new token
- Update the remote URL with new token

### Still having issues?
Try SSH method instead (see above)

---

## 🎉 Next Steps

1. **Create Personal Access Token** (5 minutes)
2. **Push to GitHub** (1 minute)
3. **Wait for build** (5-10 minutes)
4. **Download installer** (1 minute)
5. **Distribute to users!**

You're almost there! 🚀