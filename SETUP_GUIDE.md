# Deployment Guide — Portfolio + AI Recruiter Chatbot

Your site is now 3 files (plus this guide):

```
index.html          ← the full portfolio + chatbot widget (frontend)
api/chat.js          ← serverless function that talks to Groq's AI (backend)
package.json          ← lets Vercel recognize the Node function
vercel.json           ← small config for the function
```

**Why you need a host that supports "serverless functions"** (like Vercel), and not
plain GitHub Pages: GitHub Pages can only serve static files — it can't run
`api/chat.js` or keep your API key secret. Vercel's free tier does both static
hosting AND serverless functions from the same repo, so it's the easiest fit here.

---

## Step 1 — Get a free Groq API key

Groq gives free API access to Llama 3.3 70B (fast + generous free tier).

1. Go to https://console.groq.com and sign up (free).
2. Go to **API Keys** → **Create API Key**.
3. Copy the key (starts with `gsk_...`). You won't be able to see it again later.

## Step 2 — Push these files to your GitHub repo

In your existing `AmanJoshi2004/Portfolio` repo:

```bash
git checkout main
git pull
# copy in the new files: index.html (replace), api/chat.js, package.json, vercel.json
git add .
git commit -m "Add AI recruiter chatbot, new experience, and Transaction Guardian project"
git push
```

## Step 3 — Deploy to Vercel (free)

1. Go to https://vercel.com and sign up with your GitHub account.
2. Click **Add New → Project**.
3. Select your `Portfolio` repository and click **Import**.
4. Under **Environment Variables**, add:
   - Name: `GROQ_API_KEY`
   - Value: *(paste the key from Step 1)*
5. Leave all other settings as default (no framework preset needed) and click **Deploy**.
6. After ~30 seconds you'll get a live URL like `https://portfolio-yourname.vercel.app`.

That's it — the chatbot is now live and working, backed by a real LLM, with your
API key safely hidden on the server side.

### Optional: keep your github.io / custom domain
In Vercel → your project → **Settings → Domains**, add your custom domain (or point
your existing GitHub Pages domain here) if you don't want to use the `.vercel.app` URL.

---

## Step 4 — Fill in the placeholders I couldn't verify

Search the code for `UPDATE:` comments — they mark spots that need your info:

| Location | What to update |
|---|---|
| `index.html` → `.nav-cv` / hero "Download CV" | Confirm your resume PDF link is current |
| `index.html` → Transaction Guardian project card | Add the real GitHub repo link and live demo link (currently `#` placeholders) |
| `index.html` → Experience → "Accessibility & Digital Documents Assistant" | Confirm the exact job title and start date at the University of Illinois System — I inferred this from your About section text, so please double-check it |
| `index.html` → `.photo-avatar` in the About section | Swap the "AJ" initials div for your real photo: `<img src="your-photo.jpg" class="photo-avatar" style="object-fit:cover;">` |
| Sentiment Intelligence / PCB projects | Add real GitHub + live demo links (currently point to your GitHub profile, not the specific repos) |

## Step 5 — Test the chatbot

Try these on the live site:
- "Who is Aman?"
- "What's his qualification?"
- "What internship did he complete?"
- "What technologies does he know?"
- "What's the capital of France?" → should get the polite "not relevant" redirect

---

## How the guardrail works

The system prompt in `api/chat.js` instructs the model to answer only from Aman's
real resume data, and to respond with a fixed "That's not a relevant question..."
message for anything off-topic. The frontend detects that fixed sentence and styles
it slightly differently (italic) so it's visually clear it's a redirect, not an answer.

## Cost

Groq's free tier is generous (rate-limited per minute, not a hard monthly cap at time
of writing) and plenty for a portfolio site. If you ever hit limits, check
https://console.groq.com for current free-tier details, or swap the `MODEL` /
`GROQ_URL` constants in `api/chat.js` for another OpenAI-compatible provider.

## Updating the chatbot's knowledge later

Everything the bot knows lives in the `RESUME_CONTEXT` constant at the top of
`api/chat.js`. When you update your resume or add a project, update that block too —
the bot has no other source of truth about you.
