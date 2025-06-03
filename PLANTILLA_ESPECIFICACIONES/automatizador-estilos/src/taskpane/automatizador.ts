export async function aplicarEstilos(): Promise<void> {
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    for (const paragraph of paragraphs.items) {
      const text = paragraph.text.trim();

      if (/^\d+\.\d+$/.test(text)) {
        paragraph.style = "Título 2";
      } else if (/^Figura\s+\d+\.\d+/.test(text)) {
        paragraph.style = "Figura";
      } else if (/^Tabla\s+\d+\.\d+/.test(text)) {
        paragraph.style = "Tabla";
      } else if (/^Ecuación\s+\d+\.\d+/.test(text)) {
        paragraph.style = "Ecuación";
      } else if (/^Anexo\s+[A-Z]/.test(text)) {
        paragraph.style = "Anexo";
      }
    }

    await context.sync();
    console.log("Estilos aplicados correctamente.");
  });
}
