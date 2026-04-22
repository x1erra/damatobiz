'use client';

import { useEffect, useRef } from 'react';

const projects = [
  {
    name: 'APP Look-Thru Reporting',
    href: '/calc',
    meta: 'Production Streamlit tool',
    description:
      'Portfolio reporting workflow for APP statements, support files, look-through calculations, and export-ready reporting.',
  },
  {
    name: 'CI Stats Analyzer',
    href: '/csanalyzer',
    meta: 'Next.js analytics tool',
    description:
      'Custom statement generation analytics with upload management, dashboard views, and problem spotting for operational review.',
  },
];

export default function Home() {
  const menuRef = useRef<HTMLButtonElement>(null);
  const navRef = useRef<HTMLUListElement>(null);

  useEffect(() => {
    // Set year
    const yearEl = document.getElementById('year');
    if (yearEl) yearEl.textContent = String(new Date().getFullYear());

    // Menu toggle
    const menuToggle = menuRef.current;
    const navLinks = navRef.current;
    const navAnchors = navLinks?.querySelectorAll('.nav-links a');

    if (menuToggle && navLinks) {
      const handleClick = () => {
        const isOpen = navLinks.classList.toggle('open');
        menuToggle.setAttribute('aria-expanded', String(isOpen));
      };
      const handleAnchorClick = () => {
        navLinks.classList.remove('open');
        menuToggle.setAttribute('aria-expanded', 'false');
      };

      menuToggle.addEventListener('click', handleClick);

      navAnchors?.forEach((anchor) => {
        anchor.addEventListener('click', handleAnchorClick);
      });

      return () => {
        menuToggle.removeEventListener('click', handleClick);
        navAnchors?.forEach((anchor) => {
          anchor.removeEventListener('click', handleAnchorClick);
        });
      };
    }
  }, []);

  useEffect(() => {
    const faders = document.querySelectorAll('.fade-in');
    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            entry.target.classList.add('visible');
            observer.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.16, rootMargin: '0px 0px -50px 0px' }
    );
    faders.forEach((el) => observer.observe(el));
    return () => observer.disconnect();
  }, []);

  return (
    <>
      <div className="background-grid" aria-hidden="true"></div>

      <header className="site-header">
        <nav className="nav container">
          <a href="#top" className="brand">damato.biz</a>
          <button
            ref={menuRef}
            className="menu-toggle"
            aria-expanded="false"
            aria-label="Toggle navigation menu"
          >
            <span></span>
            <span></span>
          </button>
          <ul ref={navRef} className="nav-links">
            <li><a href="#about">About</a></li>
            <li><a href="#projects">Projects</a></li>
            <li><a href="#skills">Skills</a></li>
            <li><a href="#contact">Contact</a></li>
          </ul>
        </nav>
      </header>

      <main id="top">
        <section className="hero container fade-in">
          <p className="eyebrow">Software Portfolio</p>
          <h1>damato.biz</h1>
          <p className="tagline">
            A professional home for focused finance, operations, and reporting tools built with clean architecture and reliable deployment workflows.
          </p>
          <a href="#projects" className="cta">View Projects</a>
        </section>

        <section id="about" className="section container fade-in">
          <h2>About</h2>
          <p>
            I build practical software for workflows where accuracy, clarity, and maintainability matter.
            This site collects the tools I ship, the systems behind them, and the technical profile behind future work.
          </p>
        </section>

        <section id="projects" className="section container fade-in">
          <h2>Projects</h2>
          <div className="project-grid">
            {projects.map((project) => {
              const isExternal = project.href.startsWith('http');

              return (
                <article className="card project-card" key={project.name}>
                  <p className="project-meta">{project.meta}</p>
                  <h3>{project.name}</h3>
                  <p>{project.description}</p>
                  <a
                    href={project.href}
                    target={isExternal ? '_blank' : undefined}
                    rel={isExternal ? 'noopener noreferrer' : undefined}
                  >
                    Open Project
                  </a>
                </article>
              );
            })}
          </div>
        </section>

        <section id="skills" className="section container fade-in">
          <h2>Skills &amp; Tech Stack</h2>
          <div className="skills-grid" role="list" aria-label="Technology skills">
            <div className="skill" role="listitem"><i className="devicon-javascript-plain"></i><span>JavaScript</span></div>
            <div className="skill" role="listitem"><i className="devicon-typescript-plain"></i><span>TypeScript</span></div>
            <div className="skill" role="listitem"><i className="devicon-react-original"></i><span>React</span></div>
            <div className="skill" role="listitem"><i className="devicon-nextjs-original"></i><span>Next.js</span></div>
            <div className="skill" role="listitem"><i className="devicon-nodejs-plain"></i><span>Node.js</span></div>
            <div className="skill" role="listitem"><i className="devicon-express-original"></i><span>Express</span></div>
            <div className="skill" role="listitem"><i className="devicon-postgresql-plain"></i><span>PostgreSQL</span></div>
            <div className="skill" role="listitem"><i className="devicon-supabase-plain"></i><span>Supabase</span></div>
            <div className="skill" role="listitem"><i className="devicon-tailwindcss-plain"></i><span>Tailwind CSS</span></div>
            <div className="skill" role="listitem"><i className="devicon-docker-plain"></i><span>Docker</span></div>
            <div className="skill" role="listitem"><i className="devicon-githubactions-plain"></i><span>GitHub Actions</span></div>
            <div className="skill" role="listitem"><i className="devicon-vercel-original"></i><span>Vercel</span></div>
          </div>
        </section>

        <section id="contact" className="section container fade-in">
          <h2>Contact</h2>
          <p>For professional opportunities and collaborations, reach out via email.</p>
          <a className="contact-link" href="mailto:hello@damato.biz">hello@damato.biz</a>
        </section>
      </main>

      <footer className="site-footer container">
        <p>&copy; <span id="year"></span> damato.biz</p>
      </footer>
    </>
  );
}
