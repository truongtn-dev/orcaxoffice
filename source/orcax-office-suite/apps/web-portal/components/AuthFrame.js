import Link from 'next/link';

export function AuthFrame({ title, lead, children }) {
  return (
    <main className="auth-screen" id="noi-dung">
      <div className="auth-card">
        <Link className="auth-brand" href="/">
          OrcaX Office
        </Link>
        <h1>{title}</h1>
        {lead ? <p className="lead">{lead}</p> : null}
        {children}
      </div>
    </main>
  );
}
