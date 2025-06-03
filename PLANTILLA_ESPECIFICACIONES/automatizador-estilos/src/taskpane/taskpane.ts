/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import { aplicarEstilos } from "./automatizador";
import { cargarPlantilla } from "./plantilla";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "block";

    // Botón original (si aún lo usas)
    document.getElementById("run")!.onclick = () => {
      console.log("Botón Run original");
    };

    // 🗂️ Botón: Cargar plantilla
    document.getElementById("btnCargarPlantilla")!.onclick = cargarPlantilla;

    // 🎨 Botón: Aplicar estilos
    document.getElementById("btnAplicarEstilos")!.onclick = aplicarEstilos;
  }
});


export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}
