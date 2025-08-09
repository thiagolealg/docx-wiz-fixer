import { useState, useMemo } from "react";
import { useToast } from "@/hooks/use-toast";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import {
  parseDocxFile,
  normalizeParagraphs,
  buildHtmlFromParagraphs,
  checkNumbering,
  compareParagraphSets,
  htmlToDocxBlob,
  extractParagraphsFromHtml,
} from "@/utils/docx";

const downloadBlob = (blob: Blob, filename: string) => {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
};

export const DocxWorkbench = () => {
  const { toast } = useToast();

  const [primaryName, setPrimaryName] = useState<string>("");
  const [primaryParagraphs, setPrimaryParagraphs] = useState<string[]>([]);
  const [normalized, setNormalized] = useState<string[]>([]);
  const [numberingIssues, setNumberingIssues] = useState<
    { index: number; found: number; expected: number; text: string }[]
  >([]);

  const [editIndex, setEditIndex] = useState<number>(0);
  const [editText, setEditText] = useState<string>("");

  const [referenceParagraphs, setReferenceParagraphs] = useState<string[]>([]);
  const [compareResult, setCompareResult] = useState<{ missingInTarget: string[]; extraInTarget: string[] } | null>(null);

  const hasDoc = primaryParagraphs.length > 0 || normalized.length > 0;

  const displayParagraphs = useMemo(() => (normalized.length ? normalized : primaryParagraphs), [normalized, primaryParagraphs]);

  const handlePrimaryUpload = async (file: File) => {
    try {
      setPrimaryName(file.name);
      const parsed = await parseDocxFile(file);
      const baseParas = normalizeParagraphs(parsed.paragraphs);
      setPrimaryParagraphs(baseParas);
      setNormalized([]);
      setNumberingIssues([]);
      setEditIndex(0);
      setEditText("");
      toast({ title: "Documento carregado", description: `${file.name} lido com sucesso.` });
    } catch (e) {
      console.error(e);
      toast({ title: "Erro ao ler documento", description: "Verifique o arquivo .docx", variant: "destructive" });
    }
  };

  const handleReferenceUpload = async (file: File) => {
    try {
      const parsed = await parseDocxFile(file);
      const baseParas = normalizeParagraphs(parsed.paragraphs);
      setReferenceParagraphs(baseParas);
      toast({ title: "Documento de referência", description: `${file.name} carregado para comparação.` });
    } catch (e) {
      console.error(e);
      toast({ title: "Erro ao ler referência", description: "Verifique o arquivo .docx", variant: "destructive" });
    }
  };

  const runNormalization = () => {
    if (!primaryParagraphs.length) return;
    const norm = normalizeParagraphs(primaryParagraphs);
    setNormalized(norm);
    toast({ title: "Normalizado", description: "Parágrafos unificados com 1 linha entre eles." });
  };

  const runNumberingCheck = () => {
    const base = displayParagraphs;
    const issues = checkNumbering(base);
    setNumberingIssues(issues);
    if (issues.length === 0) {
      toast({ title: "Numeração OK", description: "Não foram encontrados problemas." });
    } else {
      toast({ title: "Atenção à numeração", description: `${issues.length} inconsistência(s) encontrada(s).` });
    }
  };

  const applyEdit = () => {
    const idx = Number(editIndex);
    const base = [...displayParagraphs];
    if (Number.isNaN(idx) || idx < 1 || idx > base.length) {
      toast({ title: "Índice inválido", description: "Informe um número de parágrafo existente.", variant: "destructive" });
      return;
    }
    base[idx - 1] = editText.trim();
    setNormalized(base);
    toast({ title: "Parágrafo atualizado", description: `Parágrafo ${idx} alterado.` });
  };

  const runCompare = () => {
    if (!referenceParagraphs.length || !displayParagraphs.length) {
      toast({ title: "Arquivos insuficientes", description: "Carregue o documento principal e o de referência.", variant: "destructive" });
      return;
    }
    const res = compareParagraphSets(referenceParagraphs, displayParagraphs);
    setCompareResult(res);
    toast({ title: "Comparação concluída", description: "Verifique os resultados abaixo." });
  };

  const downloadDocx = () => {
    if (!displayParagraphs.length) return;
    const html = buildHtmlFromParagraphs(displayParagraphs);
    const blob = htmlToDocxBlob(html);
    const name = primaryName ? primaryName.replace(/\.docx$/i, "-normalizado.docx") : "documento-normalizado.docx";
    downloadBlob(blob, name);
    toast({ title: "Download iniciado", description: "Arquivo .docx sendo baixado." });
  };

  return (
    <section className="w-full space-y-8">
      <Card className="shadow-[var(--shadow-elegant)]">
        <CardHeader>
          <CardTitle>Editor e Comparador de DOCX</CardTitle>
          <CardDescription>
            Faça upload do arquivo .docx, normalize parágrafos, cheque numeração, edite trechos e compare com um documento de referência.
          </CardDescription>
        </CardHeader>
        <CardContent className="space-y-6">
          <Tabs defaultValue="processar" className="w-full">
            <TabsList>
              <TabsTrigger value="processar">Processar</TabsTrigger>
              <TabsTrigger value="comparar">Comparar</TabsTrigger>
            </TabsList>
            <TabsContent value="processar" className="space-y-6">
              <div className="grid gap-4 sm:grid-cols-2">
                <div className="space-y-2">
                  <label className="text-sm">Documento principal (.docx)</label>
                  <Input type="file" accept=".docx" onChange={(e) => {
                    const f = e.target.files?.[0];
                    if (f) void handlePrimaryUpload(f);
                  }} />
                </div>
                <div className="flex items-end gap-2">
                  <Button onClick={runNormalization} variant="default">Normalizar</Button>
                  <Button onClick={runNumberingCheck} variant="secondary">Checar numeração</Button>
                  <Button onClick={downloadDocx} variant="outline">Baixar .docx</Button>
                </div>
              </div>

              {hasDoc && (
                <div className="grid gap-6 lg:grid-cols-2">
                  <div className="space-y-3">
                    <h3 className="text-lg font-medium">Pré-visualização (parágrafos)</h3>
                    <div className="h-80 overflow-auto rounded-md border p-4 bg-card">
                      <ol className="list-decimal pl-6 space-y-2">
                        {displayParagraphs.map((p, i) => (
                          <li key={i} className="text-sm leading-relaxed">
                            {p}
                          </li>
                        ))}
                      </ol>
                    </div>
                  </div>

                  <div className="space-y-3">
                    <h3 className="text-lg font-medium">Editar parágrafo específico</h3>
                    <div className="grid gap-3 sm:grid-cols-3">
                      <div className="space-y-2">
                        <label className="text-sm">Nº do parágrafo</label>
                        <Input type="number" min={1} value={editIndex} onChange={(e) => setEditIndex(Number(e.target.value))} />
                      </div>
                      <div className="sm:col-span-2 space-y-2">
                        <label className="text-sm">Novo conteúdo</label>
                        <Textarea value={editText} onChange={(e) => setEditText(e.target.value)} rows={5} />
                      </div>
                    </div>
                    <div className="flex gap-2">
                      <Button onClick={applyEdit}>Aplicar alteração</Button>
                    </div>

                    {numberingIssues.length > 0 && (
                      <div className="mt-4 rounded-md border p-3">
                        <p className="text-sm font-medium mb-2">Inconsistências de numeração</p>
                        <ul className="list-disc pl-5 space-y-1 text-sm">
                          {numberingIssues.map((n) => (
                            <li key={n.index}>Parágrafo {n.index + 1}: encontrado {n.found}, esperado {n.expected}</li>
                          ))}
                        </ul>
                      </div>
                    )}
                  </div>
                </div>
              )}
            </TabsContent>

            <TabsContent value="comparar" className="space-y-6">
              <div className="grid gap-4 sm:grid-cols-2">
                <div className="space-y-2">
                  <label className="text-sm">Documento de referência (.docx)</label>
                  <Input type="file" accept=".docx" onChange={(e) => {
                    const f = e.target.files?.[0];
                    if (f) void handleReferenceUpload(f);
                  }} />
                </div>
                <div className="flex items-end gap-2">
                  <Button onClick={runCompare} variant="default">Comparar</Button>
                </div>
              </div>

              {compareResult && (
                <div className="grid gap-6 lg:grid-cols-2">
                  <div className="space-y-2">
                    <h3 className="text-lg font-medium">Faltando no documento principal</h3>
                    <div className="h-72 overflow-auto rounded-md border p-4 bg-card">
                      {compareResult.missingInTarget.length === 0 ? (
                        <p className="text-sm text-muted-foreground">Nada faltando. Tudo ok!</p>
                      ) : (
                        <ul className="list-disc pl-5 space-y-2 text-sm">
                          {compareResult.missingInTarget.map((p, i) => (
                            <li key={i}>{p}</li>
                          ))}
                        </ul>
                      )}
                    </div>
                  </div>
                  <div className="space-y-2">
                    <h3 className="text-lg font-medium">Excessos no documento principal</h3>
                    <div className="h-72 overflow-auto rounded-md border p-4 bg-card">
                      {compareResult.extraInTarget.length === 0 ? (
                        <p className="text-sm text-muted-foreground">Nenhum excesso encontrado.</p>
                      ) : (
                        <ul className="list-disc pl-5 space-y-2 text-sm">
                          {compareResult.extraInTarget.map((p, i) => (
                            <li key={i}>{p}</li>
                          ))}
                        </ul>
                      )}
                    </div>
                  </div>
                </div>
              )}
            </TabsContent>
          </Tabs>
        </CardContent>
      </Card>
    </section>
  );
};

export default DocxWorkbench;
