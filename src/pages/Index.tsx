import { Helmet, HelmetProvider } from "react-helmet-async";
import DocxWorkbench from "@/components/docx/DocxWorkbench";

const Index = () => {
  const title = "Editor DOCX Online – Normalização, Numeração e Comparação";
  const description = "Carregue, normalize e compare arquivos .docx diretamente no navegador. Ajuste parágrafos, verifique numeração e exporte o resultado.";
  const canonical = typeof window !== "undefined" ? window.location.href : "";

  return (
    <div className="min-h-screen bg-background text-foreground">
      <HelmetProvider>
        <Helmet>
          <title>{title}</title>
          <meta name="description" content={description} />
          <link rel="canonical" href={canonical} />
          <meta property="og:title" content={title} />
          <meta property="og:description" content={description} />
          <script type="application/ld+json">{JSON.stringify({
            "@context": "https://schema.org",
            "@type": "SoftwareApplication",
            name: "Editor DOCX Online",
            description,
            applicationCategory: "BusinessApplication",
            operatingSystem: "Web",
          })}</script>
        </Helmet>
      </HelmetProvider>

      <header className="relative overflow-hidden">
        <div className="pointer-events-none absolute inset-0 hero-glow" aria-hidden="true" />
        <div className="container py-16 sm:py-20">
          <div className="mx-auto max-w-3xl text-center">
            <h1 className="text-4xl sm:text-5xl font-bold tracking-tight">
              Editor e Comparador de Arquivos DOCX
            </h1>
            <p className="mt-4 text-lg text-muted-foreground">
              Leia, normalize e compare documentos .docx, diretamente no navegador — sem instalações.
            </p>
          </div>
        </div>
      </header>

      <main className="container pb-20">
        <DocxWorkbench />
      </main>
    </div>
  );
};

export default Index;
