document.addEventListener("DOMContentLoaded", () => {

    const form = document.getElementById("ver_escala");

    if (!form) {
        console.error("Formulário não encontrado.");
        return;
    }

    form.addEventListener("submit", async (event) => {
        event.preventDefault();

        const escalaFile = document.getElementById("escala")?.files[0];
        const respostasFile = document.getElementById("resposta")?.files[0];

        if (!escalaFile || !respostasFile) {
            alert("Selecione os dois arquivos antes de enviar.");
            return;
        }

        let escala;
        let respostas;

        // Ler os dois arquivos
        escala = await lerArquivo(escalaFile, "ESCALA");
        respostas = await lerArquivo(respostasFile, "RESPOSTAS");

        var nomePorDia;
        var nomePorHorario;
        nomePorDia = getNomePorDia(escala);
        nomePorHorario = getNomePorHorario(escala);

        exibirEscala(escala, nomePorDia, nomePorHorario);
    });

});


function lerArquivo(file, nome) {
    return new Promise((resolve, reject) => {

        const reader = new FileReader();

        reader.onload = function (e) {

            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: "array" });

                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];

                var json;

                if (nome === "ESCALA") {
                    var ref = "A3:I14";
                    json = XLSX.utils.sheet_to_json(sheet, { range: ref, raw: false });
                } else {
                    const range = XLSX.utils.decode_range(sheet["!ref"]);
                    range.s.c = 1;
                    range.e.c = 3;
                    sheet["!ref"] = XLSX.utils.encode_range(range);
                    json = XLSX.utils.sheet_to_json(sheet, { range: ref, raw: false });
                }

                resolve(json);
            } catch (erro) {
                console.error(`Erro ao processar ${nome}:`, erro);
                reject(`Erro ao ler o arquivo ${nome}`);
            }

        };

        reader.onerror = function () {
            alert(`Erro ao carregar o arquivo ${nome}`);
        };

        reader.readAsArrayBuffer(file);
    });
}

function separarNomes(texto) {
    if (!texto) return [];

    return texto
        .split(/\s*(?:,|\be\b)\s*/i)
        .map(n => n.trim())
        .filter(n => n.length > 0);
}

function getNomePorDia(escala) {

    const nomePorDia = {};

    for (let i = 0; i < escala.length; i++) {

        nomePorDia[escala[i]["SÁBADO"]] = separarNomes(escala[i]["17H"]);

        nomePorDia[escala[i]["DOMINGO"]] = [
            ...separarNomes(escala[i]["7H"]),
            ...separarNomes(escala[i]["9H"]),
            ...separarNomes(escala[i]["11H"]),
            ...separarNomes(escala[i]["17H_1"]),
            ...separarNomes(escala[i]["19H"])
        ];
    }

    console.log(nomePorDia);
    return nomePorDia;
}

function getNomePorHorario(escala) {
    const nomePorHorario = { "17H": [], "7H": [], "9H": [], "11H": [], "17H_1": [], "19H": [] };

    for (let i = 0; i < escala.length; i++) {
        nomePorHorario["17H"].push(...separarNomes(escala[i]["17H"]));
        nomePorHorario["7H"].push(...separarNomes(escala[i]["7H"]));
        nomePorHorario["9H"].push(...separarNomes(escala[i]["9H"]));
        nomePorHorario["11H"].push(...separarNomes(escala[i]["11H"]));
        nomePorHorario["17H_1"].push(...separarNomes(escala[i]["17H_1"]));
        nomePorHorario["19H"].push(...separarNomes(escala[i]["19H"]));
    }
    console.log(nomePorHorario);
    return nomePorHorario;


}

function verificarRespostas(nomes) {
    var status = ["CERTO", "ERRADO", "ATENCAO"]

    var retorno = {};

    for (let i = 0; i < nomes.length; i++) {
        retorno[nomes[i]] = status[Math.floor(Math.random() * status.length)];
    }
    return retorno;
}

function adicionarCor(coluna, ver) {
    var cores = {
        "CERTO": "green",
        "ERRADO": "red",
        "ATENCAO": "gold"
    }
    for (let nome in ver) {
        var span = document.createElement("span");
        span.style.color = cores[ver[nome]];
        span.textContent = nome + " ";
        coluna.appendChild(span);
    }
    const spans = coluna.querySelectorAll("span");
    spans.forEach((span, index) => {
        if (index < spans.length - 1) {
            span.insertAdjacentText("afterend", " - ");
        }
    });
}

function adicionarLinhaNaTabelaComCor(linha) {
    var diaSab = linha["SÁBADO"];
    var diaDom = linha["DOMINGO"];

    var missaSab17 = separarNomes(linha["17H"]);
    var missaSab17ver = verificarRespostas(missaSab17);
    var missaDom7 = separarNomes(linha["7H"]);
    var missaDom7ver = verificarRespostas(missaDom7);
    var missaDom9 = separarNomes(linha["9H"]);
    var missaDom9ver = verificarRespostas(missaDom9);
    var missaDom11 = separarNomes(linha["11H"]);
    var missaDom11ver = verificarRespostas(missaDom11);
    var missaDom17 = separarNomes(linha["17H_1"]);
    var missaDom17ver = verificarRespostas(missaDom17);
    var missaDom19 = separarNomes(linha["19H"]);
    var missaDom19ver = verificarRespostas(missaDom19);

    var linhaTabela = document.createElement("tr");
    for (let i = 0; i < 8; i++) {
        var col = document.createElement("td");
        if (i == 0) {
            col.textContent = diaSab;
        } else if (i == 1) {
            adicionarCor(col, missaSab17ver)
        } else if (i == 2) {
            col.textContent = diaDom;
        } else if (i == 3) {
            adicionarCor(col, missaDom7ver)
        } else if (i == 4) {
            adicionarCor(col, missaDom9ver)
        } else if (i == 5) {
            adicionarCor(col, missaDom11ver)
        } else if (i == 6) {
            adicionarCor(col, missaDom17ver)
        } else if (i == 7) {
            adicionarCor(col, missaDom19ver)
        }
        linhaTabela.appendChild(col);
    }
    return linhaTabela;
}

function exibirEscala(escala, nomePorDia) {
    const localExibirEscala = document.getElementById("exibirEscala");
    localExibirEscala.innerHTML = "";
    const escalaCabecalho = document.createElement("thead");
    const cab = document.createElement("tr");
    const escalaCorpo = document.createElement("tbody");
    let cabecalho = ["SÁBADO", "17H", "DOMINGO", "7H", "9H", "11H", "17H", "19H"];
    for (let c in cabecalho) {
        var col = document.createElement("th");
        col.textContent = cabecalho[c];
        cab.appendChild(col);
    }
    escalaCabecalho.appendChild(cab);
    localExibirEscala.appendChild(escalaCabecalho);
    cabecalho[6] = "17H_1";
    for (let i = 0; i < escala.length; i++) {
        var linha = adicionarLinhaNaTabelaComCor(escala[i]);
        escalaCorpo.appendChild(linha);

    }
    localExibirEscala.appendChild(escalaCorpo);
}