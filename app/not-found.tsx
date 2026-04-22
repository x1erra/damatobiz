import Link from "next/link";

export default function NotFound() {
  return (
    <>
      <div className="background-grid" aria-hidden="true"></div>
      <main className="not-found">
        <section className="not-found-card fade-in visible">
          <p className="eyebrow">404 Error</p>
          <h1>Page Not Found</h1>
          <p>The page you requested does not exist or may have moved.</p>
          <Link className="cta" href="/">Return Home</Link>
        </section>
      </main>
    </>
  );
}
