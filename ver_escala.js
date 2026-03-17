document.addEventListener("DOMContentLoaded", () => {

    const form = document.getElementById("ver_escala");

    if (!form) {
        console.error("Formulário não encontrado.");
        return;
    }

    form.addEventListener("submit", (event) => {
        event.preventDefault();

        const escalaFile = document.getElementById("escala")?.files[0];
        const respostasFile = document.getElementById("resposta")?.files[0];

        if (!escalaFile || !respostasFile) {
            alert("Selecione os dois arquivos antes de enviar.");
            return;
        }

        // Ler os dois arquivos
        lerArquivo(escalaFile, "ESCALA");
        lerArquivo(respostasFile, "RESPOSTAS");
    });

});


function lerArquivo(file, nome) {

    const reader = new FileReader();

    reader.onload = function (e) {

        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });

            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            const json = XLSX.utils.sheet_to_json(sheet);

            console.log(`===== ${nome} =====`);
            console.table(json);

        } catch (erro) {
            console.error(`Erro ao processar ${nome}:`, erro);
            alert(`Erro ao ler o arquivo ${nome}`);
        }

    };

    reader.onerror = function () {
        alert(`Erro ao carregar o arquivo ${nome}`);
    };

    reader.readAsArrayBuffer(file);
}