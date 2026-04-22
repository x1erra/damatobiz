# damato.biz

Professional portfolio and tool catalog for Damato-built software.

## Repository Layout

- `app/` - Next.js portfolio site and App Router routes.
- `app/calc/` - Redirect route so `damato.biz/calc` opens the APP Look-Thru Reporting deployment at `calc.damato.biz`.
- `app/(csanalyzer)/csanalyzer/` - CI Stats Analyzer route.
- `components/`, `lib/`, `types/` - Shared code used by the Next.js tools.
- `tools/app-statements-calculator/` - APP Look-Thru Reporting Streamlit app copied from `x1erra/APPStatementsCalculator` without changing the calculator implementation.

## Local Development

Run the portfolio and Next.js tools:

```bash
npm install
npm run dev
```

Open `http://localhost:3000`.

Run the APP Look-Thru Reporting tool:

```bash
cd tools/app-statements-calculator
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Deployment

The repo root is the Vercel project for `damato.biz`.

The Streamlit calculator remains deployable from `tools/app-statements-calculator` with its existing `Dockerfile`, `docker-compose.yml`, and `render.yaml`. Keep `calc.damato.biz` pointed at that app unless the calculator is intentionally rewritten as a Next.js route later.

## Branches

Use `master` as the live branch for the consolidated repo. The older `main` branch exists remotely and can be archived or deleted after Vercel is configured to deploy only `master`.

## Vercel Cleanup

Recommended target state:

- Keep one Vercel project for this repo root: `damatobiz`.
- Set the production branch to `master`.
- Keep `damato.biz` assigned to that project.
- Remove the duplicate `damatobiz1` project after the `damatobiz` deployment is healthy.
- Keep `calc.damato.biz` pointed at the Streamlit calculator deployment unless the calculator is deliberately rebuilt as a Next.js route.
