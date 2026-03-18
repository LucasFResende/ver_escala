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
        const conversaoFile = document.getElementById("conversao")?.files[0];

        if (!escalaFile || !respostasFile) {
            alert("Selecione os dois arquivos antes de enviar.");
            return;
        }

        let escala;
        let respostas;
        let conversao;

        // Ler os dois arquivos
        escala = await lerArquivo(escalaFile, "ESCALA");
        respostas = await lerArquivo(respostasFile, "RESPOSTAS");
        conversao = await lerArquivo(conversaoFile, "CONVERSAO");


        respostas = ajeitarRespostas(respostas, conversao);


        // avaliarEscala(escala, respostas);

        const nomePorDia = getNomePorDia(escala);
        const nomePorHorario = getNomePorHorario(escala);
        exibirEscala(escala, respostas, conversao);
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
                } else if (nome === "RESPOSTAS") {
                    const range = XLSX.utils.decode_range(sheet["!ref"]);
                    range.s.c = 1;
                    range.e.c = 3;
                    sheet["!ref"] = XLSX.utils.encode_range(range);
                    json = XLSX.utils.sheet_to_json(sheet, { range: ref, raw: false });
                    for (let x in json) {
                        const { ["Nome do acólito ( nome e sobrenome ) "]: nomeAcolito,
                            ["Dias"]: dias,
                            ...resto } = json[x];
                        json[x] = { nome: nomeAcolito, restricao: dias, horario: Object.values(resto)[0] };
                    }
                } else if (nome === "CONVERSAO") {
                    json = XLSX.utils.sheet_to_json(sheet, { raw: false });
                    console.log(json);
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

function ajeitarRespostas(respostas, conversao) {
    respostas = respostas.filter(r =>
        r.horario !== "NÃO VOU SERVIR ESSE MÊS"
    );

    for (let i = 0; i < respostas.length; i++) {
        if (respostas[i].restricao === undefined) {
            respostas[i].restricao = [];
        } else {
            respostas[i].restricao = respostas[i].restricao.split(",");
        }
        respostas[i].horario = respostas[i].horario.split(",");
        for (let j = 0; j < respostas[i].horario.length; j++) {
            if (respostas[i].horario[j].trim() === "SABADOS AS 17H") respostas[i].horario[j] = "17H";
            if (respostas[i].horario[j].trim() === "DOMINGOS AS 7H") respostas[i].horario[j] = "7H";
            if (respostas[i].horario[j].trim() === "DOMINGOS AS 9H ( Max 4 acolitos )") respostas[i].horario[j] = "9H";
            if (respostas[i].horario[j].trim() === "DOMINGOS AS 11H") respostas[i].horario[j] = "11H";
            if (respostas[i].horario[j].trim() === "DOMINGOS AS 17h" || respostas[i].horario[j].trim() === "DOMINGOS AS 17H") respostas[i].horario[j] = "17H_1";
            if (respostas[i].horario[j].trim() === "DOMINGOS AS 19H") respostas[i].horario[j] = "19H";
            if (respostas[i].horario[j].trim() === "SEM PREFERÊNCIAS") respostas[i].horario = ["17H", "7H", "9H", "11H", "17H_1", "19H"];
        }
        for (let j = 0; j < conversao.length; j++) {
            if (conversao[j].NomeResposta.trim() === respostas[i].nome.trim()) {
                respostas[i].nome = conversao[j].NomeEscala.trim();
            }
        }
    }
    console.log(respostas);
    return respostas;
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
    return nomePorHorario;


}

function verificarRespostas(nomes, dia, hor, respostas) {
    var status = ["CERTO", "ERRADO", "ATENCAO"]

    var retorno = {};

    for (let i = 0; i < nomes.length; i++) {
        var achou = false;
        for (let j = 0; j < respostas.length; j++) {
            if (respostas[j].nome.trim() === nomes[i]) {
                achou = true;
                if (respostas[j].restricao.includes(dia)) {
                    retorno[nomes[i]] = status[1];
                } else if (!respostas[j].horario.includes(hor)) {
                    retorno[nomes[i]] = status[1];
                } else {
                    retorno[nomes[i]] = status[0];
                }
            }
        }
        if (!achou) {
            retorno[nomes[i]] = status[2];
        }
    }
    // console.log(retorno);
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

function adicionarLinhaNaTabelaComCor(linha, respostas) {
    var diaSab = linha["SÁBADO"];
    var diaDom = linha["DOMINGO"];

    var missaSab17 = separarNomes(linha["17H"]);
    var missaSab17ver = verificarRespostas(missaSab17, diaSab, "17H", respostas);
    var missaDom7 = separarNomes(linha["7H"]);
    var missaDom7ver = verificarRespostas(missaDom7, diaDom, "7H", respostas);
    var missaDom9 = separarNomes(linha["9H"]);
    var missaDom9ver = verificarRespostas(missaDom9, diaDom, "9H", respostas);
    var missaDom11 = separarNomes(linha["11H"]);
    var missaDom11ver = verificarRespostas(missaDom11, diaDom, "11H", respostas);
    var missaDom17 = separarNomes(linha["17H_1"]);
    var missaDom17ver = verificarRespostas(missaDom17, diaDom, "17H_1", respostas);
    var missaDom19 = separarNomes(linha["19H"]);
    var missaDom19ver = verificarRespostas(missaDom19, diaDom, "19H", respostas);

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

function exibirEscala(escala, respostas) {
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
        var linha = adicionarLinhaNaTabelaComCor(escala[i], respostas);
        escalaCorpo.appendChild(linha);

    }
    localExibirEscala.appendChild(escalaCorpo);
}

function avaliarEscala(escala, respostas) {
    var nomePorDia;
    var nomePorHorario;
    nomePorDia = getNomePorDia(escala);
    nomePorHorario = getNomePorHorario(escala);
    var avaliacao = {}
    console.log(escala);
    console.log(nomePorDia);
    console.log(nomePorHorario);


}