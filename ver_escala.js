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
        var nomes = [];
        for (let x in nomePorHorario) {
            nomes.push(...nomePorHorario[x]);
        }
        exibirEscala(escala, respostas, nomes);
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
                        json[x] = { nome: nomeAcolito.trim(), restricao: dias, horario: Object.values(resto)[0] };
                    }
                } else if (nome === "CONVERSAO") {
                    json = XLSX.utils.sheet_to_json(sheet, { raw: false });
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

function verificarNaoEscalado(respostas, nomes) {
    var observacao = {}
    for (let i = 0; i < respostas.length; i++) {
        if (!nomes.includes(respostas[i].nome)) {
            observacao[respostas[i].nome] = {
                cor: "red",
                observacao: "Não escalado"
            };
        }
    }
    return observacao
}

function verificarRespostas(nomes, dia, hor, respostas, observacao) {
    var status = ["CERTO", "ERRADO", "ATENCAO"]
    var retorno = {};

    for (let i = 0; i < nomes.length; i++) {
        var achou = false;
        for (let j = 0; j < respostas.length; j++) {
            if (respostas[j].nome.trim() === nomes[i]) {
                achou = true;
                if (respostas[j].restricao.includes(dia)) {
                    retorno[nomes[i]] = status[1];
                    observacao[nomes[i]] = {
                        cor: "red",
                        observacao: "Restrição no dia " + dia
                    }
                } else if (!respostas[j].horario.includes(hor)) {
                    retorno[nomes[i]] = status[1];
                    observacao[nomes[i]] = {
                        cor: "red",
                        observacao: "Restrição no horário " + hor
                    }
                } else {
                    retorno[nomes[i]] = status[0];
                    if (nomes[i] in observacao) delete observacao[nomes[i]];
                }
            }
        }
        if (!achou) {
            retorno[nomes[i]] = status[2];
            observacao[nomes[i]] = {
                cor: "gold",
                observacao: "Não respondeu ao formulário"
            }
        }
    }

    return retorno;
}

function formatarClasse(nome) {
    return nome
        .normalize("NFD")                // remove acento
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, "_")            // espaço → _
        .replace(/[^\w-]/g, "")          // remove resto
        .toLowerCase();
}

function adicionarCor(coluna, ver) {
    var cores = {
        "CERTO": "green",
        "ERRADO": "red",
        "ATENCAO": "gold"
    }
    for (let nome in ver) {
        var span = document.createElement("span");
        span.classList.add(formatarClasse(nome));
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

function adicionarLinhaNaTabelaComCor(linha, respostas, index, observacao) {
    var diaSab = linha["SÁBADO"];
    var diaDom = linha["DOMINGO"];

    var missaSab17 = separarNomes(linha["17H"]);
    var missaSab17ver = verificarRespostas(missaSab17, diaSab, "17H", respostas, observacao);
    var missaDom7 = separarNomes(linha["7H"]);
    var missaDom7ver = verificarRespostas(missaDom7, diaDom, "7H", respostas, observacao);
    var missaDom9 = separarNomes(linha["9H"]);
    var missaDom9ver = verificarRespostas(missaDom9, diaDom, "9H", respostas, observacao);
    var missaDom11 = separarNomes(linha["11H"]);
    var missaDom11ver = verificarRespostas(missaDom11, diaDom, "11H", respostas, observacao);
    var missaDom17 = separarNomes(linha["17H_1"]);
    var missaDom17ver = verificarRespostas(missaDom17, diaDom, "17H_1", respostas, observacao);
    var missaDom19 = separarNomes(linha["19H"]);
    var missaDom19ver = verificarRespostas(missaDom19, diaDom, "19H", respostas, observacao);

    var linhaTabela = document.createElement("tr");
    linhaTabela.id = "semana" + index;

    for (let i = 0; i < 8; i++) {
        var col = document.createElement("td");

        if (i == 0) {
            col.textContent = diaSab;
        } else if (i == 1) {
            col.classList.add("17H");
            adicionarCor(col, missaSab17ver);
        } else if (i == 2) {
            col.textContent = diaDom;
        } else if (i == 3) {
            col.classList.add("7H");
            adicionarCor(col, missaDom7ver);
        } else if (i == 4) {
            col.classList.add("9H");
            adicionarCor(col, missaDom9ver);
        } else if (i == 5) {
            col.classList.add("11H");
            adicionarCor(col, missaDom11ver);
        } else if (i == 6) {
            col.classList.add("17H_1");
            adicionarCor(col, missaDom17ver);
        } else if (i == 7) {
            col.classList.add("19H");
            adicionarCor(col, missaDom19ver);
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
    var observacao = {};

    for (let c in cabecalho) {
        var col = document.createElement("th");
        col.textContent = cabecalho[c];
        cab.appendChild(col);
    }
    escalaCabecalho.appendChild(cab);
    localExibirEscala.appendChild(escalaCabecalho);
    cabecalho[6] = "17H_1";
    for (let i = 0; i < escala.length; i++) {
        var linha = adicionarLinhaNaTabelaComCor(escala[i], respostas, i, observacao);
        escalaCorpo.appendChild(linha);
    }
    localExibirEscala.appendChild(escalaCorpo);
    for (let i = 0; i < escala.length; i++) {
        verificarQuantidadeDeVezesEscaladoPorSemana(observacao, i);
    }
    verificarQuantidadeDeVezesEscaladoGeral(observacao, respostas);
    exibirObservacoes(observacao);
}

function verificarQuantidadeDeVezesEscaladoGeral(observacao, respostas) {
    for (let i = 0; i < respostas.length; i++) {
        var classe = `.${formatarClasse(respostas[i].nome)}`;
        if (Object.keys(observacao).includes(respostas[i].nome)) {
            continue;
        }
        if (document.querySelectorAll(classe).length > 2) {
            observacao[respostas[i].nome] = {
                cor: "gold",
                observacao: "Escalado mais de duas vezes"
            }
            const elementos = document.querySelectorAll(classe);

            elementos.forEach(el => {
                el.style.color = "gold";
            });
        } else if (document.querySelectorAll(classe).length < 1) {
            observacao[respostas[i].nome] = {
                cor: "red",
                observacao: "Não escalado"
            }
        }
    }
}

function verificarQuantidadeDeVezesEscaladoPorSemana(observacao, index) {

    const linha = document.getElementById("semana" + index);
    const spans = linha.querySelectorAll("td span");

    // Extrair os nomes de cada span
    const nomes = Array.from(spans).map(s => s.textContent.trim());

    // Encontrar duplicados
    const duplicados = nomes.filter((nome, i) => nomes.indexOf(nome) !== i);

    // Remover repetidos
    const duplicadosUnicos = [...new Set(duplicados)];


    duplicadosUnicos.forEach(nome => {
        // Verifica se o nome já está na observação
        if (!(nome in observacao)) {
            observacao[nome] = {
                cor: "gold",
                observacao: "Escalado mais de uma vez nesta semana"
            };

            // Alterar cor no DOM
            spans.forEach(s => {
                if (s.textContent.trim() === nome) {
                    s.style.color = "gold";
                }
            });
        }
    });
}

function exibirObservacoes(observacao) {
    const localExibirEscala = document.getElementById("observacao");
    localExibirEscala.innerHTML = '<h1 style="color: white;">Observações</h1>';
    let index = 0;
    for (let o in observacao) {
        const id = `collapse${index}`
        localExibirEscala.innerHTML += `<div class="accordion">
            <div class="accordion-item">
                <h2 class="accordion-header">
                    <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse"
                        data-bs-target="#${id}" aria-expanded="true" aria-controls="collapseOne" style="color: ${observacao[o].cor}">
                        ${o}
                    </button>
                </h2>
                <div id="${id}" class="accordion-collapse collapse" data-bs-parent="#accordionExample">
                    <div class="accordion-body">
                        <h2>${observacao[o].observacao}</h2>
                    </div>
                </div>
            </div>
        </div>`;
        index++;
    }
}