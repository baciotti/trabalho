/*****************************************************************
 *
 * HEIDEGGER
 * Normalizador de Relatórios
 *
 *****************************************************************/

class Heidegger {

    constructor(){

        this.dropZone=document.getElementById("dropZone");
        this.fileInput=document.getElementById("fileInput");

        this.status=document.getElementById("statusMessage");
        this.progress=document.getElementById("progressBar");
        this.history=document.getElementById("history");

        this.events();

        this.setStatus("Aguardando arquivo...",0);

    }

    /************************************************************/

    events(){

        this.dropZone.onclick=()=>this.fileInput.click();

        this.fileInput.onchange=(e)=>{

            if(e.target.files.length){

                this.process(e.target.files);

            }

            this.fileInput.value="";

        };

        this.dropZone.ondragover=(e)=>{

            e.preventDefault();

            this.dropZone.classList.add("drag");

        };

        this.dropZone.ondragleave=()=>{

            this.dropZone.classList.remove("drag");

        };

        this.dropZone.ondrop=(e)=>{

            e.preventDefault();

            this.dropZone.classList.remove("drag");

            this.process(e.dataTransfer.files);

        };

    }

    /************************************************************/

    async process(files){

        for(const file of files){

            try{

                await this.processFile(file);

            }
            catch(err){

                console.error(err);

                this.setStatus(err.message,0);

            }

        }

        this.setStatus("Aguardando arquivo...",0);

    }

    /************************************************************/

    async processFile(file){

        if(!file.name.toLowerCase().endsWith(".csv.gz")){

            throw new Error("Arquivo inválido.");

        }

        this.setStatus("Descompactando...",10);

        const text=await this.readFile(file);

        this.setStatus("Lendo relatório...",30);

        const rows=this.parseCSV(text);
		
		if(rows.length < 6){

		throw new Error("Relatório inválido.");

		}

        const meta=this.extractMetadata(rows,file.name);

        this.setStatus("Preparando dados...",60);

        const table=this.prepareTable(rows);

        this.setStatus("Gerando Excel...",80);

        const wb=this.createWorkbook(meta,table);

        this.setStatus("Download...",95);

        this.download(wb,meta);

        this.addHistory(meta);

        this.setStatus("Conversão concluída.",100);

        await this.sleep(800);

    }

    /************************************************************/

    setStatus(text,percent){

        this.status.textContent=text;

        this.progress.style.width=percent+"%";

    }

    /************************************************************/

    addHistory(meta){

        const empty=this.history.querySelector(".history-empty");

        if(empty) empty.remove();

        const div=document.createElement("div");

        div.className="history-item";

        div.innerHTML=`

            <strong>${meta.cpf}</strong>

            <small>

                ${meta.totalEventos} eventos

            </small>

        `;

        this.history.prepend(div);

    }

    /************************************************************/

    sleep(ms){

        return new Promise(r=>setTimeout(r,ms));

    }



    /************************************************************
     * Lê e descompacta o arquivo .csv.gz
     ************************************************************/
    async readFile(file){

        const buffer = await file.arrayBuffer();

        const bytes = new Uint8Array(buffer);

        // Arquivo .gz
        const inflated = fflate.gunzipSync(bytes);

        return this.decode(inflated);

    }

    /************************************************************
     * Decodificação
     ************************************************************/
    decode(bytes){

        // Tenta UTF-8 primeiro
        let text = new TextDecoder("utf-8").decode(bytes);

        // Se encontrou caracteres inválidos,
        // tenta Latin1
        if(text.includes("�")){

            text = new TextDecoder("latin1").decode(bytes);

        }

        return text;

    }

    /************************************************************
     * Converte CSV em Array
     ************************************************************/
    parseCSV(text){

        const lines = text
            .replace(/\r/g,"")
            .split("\n")
            .filter(l => l.trim() !== "");

        return lines.map(line => this.parseLine(line));

    }

    /************************************************************
     * Parser de uma linha CSV
     ************************************************************/
    parseLine(line){

        const row = [];

        let value = "";

        let quoted = false;

        for(let i=0;i<line.length;i++){

            const c = line[i];

            if(c === '"'){

                // Aspas escapadas ("")
                if(quoted && line[i+1] === '"'){

                    value += '"';
                    i++;
                    continue;

                }

                quoted = !quoted;
                continue;

            }

            if(c === "," && !quoted){

                row.push(value);
                value = "";
                continue;

            }

            value += c;

        }

        row.push(value);

        return row;

    }


    /************************************************************
     * Extrai os metadados do relatório
     ************************************************************/
    extractMetadata(rows, fileName){

        const cpf = (rows[1]?.[1] || "").replace(/\D/g,"");

        const dataInicial = rows[2]?.[1] || "";

        const dataFinal = rows[3]?.[1] || "";

        const totalEventos = Math.max(rows.length - 6, 0);

        return {

            cpf: this.formatCPF(cpf),

            cpfOriginal: cpf,

            dataInicial,

            dataFinal,

            totalEventos,

            arquivoOriginal: fileName,

            dataConversao: new Date()

        };

    }

    /************************************************************
     * Mantém apenas a tabela de eventos
     ************************************************************/
    prepareTable(rows){

        // Remove as cinco primeiras linhas
        let table = rows.slice(5);
		
		// Remove as colunas que não precisamos
		const COLUNAS_REMOVER = [0, 4, 5, 9];

		table = table.map(row =>
			row.filter((_, index) => !COLUNAS_REMOVER.includes(index))
		);

		// Remove o ID entre colchetes da coluna "Info. Adicional"
		table = table.map(row => {

			if(row[3]){

				row[3] = row[3].replace(/^\[\d+\]\s*/, "");

			}

			return row;

		});
		

        // Remove linhas totalmente vazias
        table = table.filter(row =>
            row.some(col => String(col).trim() !== "")
        );

        return table;

    }

    /************************************************************
     * Formata CPF
     ************************************************************/
    formatCPF(cpf){

        cpf = String(cpf).replace(/\D/g,"");

        if(cpf.length !== 11){

            return cpf;

        }

        return cpf.replace(

            /(\d{3})(\d{3})(\d{3})(\d{2})/,

            "$1.$2.$3-$4"

        );

    }
	
	
	    /************************************************************
     * Gera o Workbook
     ************************************************************/
    createWorkbook(meta, table){

        const wb = XLSX.utils.book_new();

        /******************************************************
         * Aba Resumo
         ******************************************************/

        const resumo = [

            ["Campo","Valor"],

            ["CPF", meta.cpf],

            ["Data Inicial", meta.dataInicial],

            ["Data Final", meta.dataFinal],

            ["Quantidade de Eventos", meta.totalEventos],

            ["Arquivo Original", meta.arquivoOriginal],

            ["Data da Conversão", meta.dataConversao.toLocaleString("pt-BR")]

        ];

        const wsResumo = XLSX.utils.aoa_to_sheet(resumo);

        XLSX.utils.book_append_sheet(

            wb,

            wsResumo,

            "Resumo"

        );

        /******************************************************
         * Aba Eventos
         ******************************************************/

        const wsEventos = XLSX.utils.aoa_to_sheet(table);

        if(wsEventos["!ref"]){

            wsEventos["!autofilter"] = {

                ref: wsEventos["!ref"]

            };

        }

        wsEventos["!cols"] = this.calculateColumnWidths(table);

        XLSX.utils.book_append_sheet(

            wb,

            wsEventos,

            "Eventos"

        );

        return wb;

    }

    /************************************************************
     * Calcula largura das colunas
     ************************************************************/
    calculateColumnWidths(rows){

        const widths = [];

        rows.forEach(row=>{

            row.forEach((cell,index)=>{

                const len = String(cell ?? "").length;

                if(!widths[index]){

                    widths[index]={wch:len};

                }
                else{

                    widths[index].wch = Math.max(

                        widths[index].wch,

                        len

                    );

                }

            });

        });

        widths.forEach(col=>{

            col.wch = Math.max(

                10,

                Math.min(col.wch+2,60)

            );

        });

        return widths;

    }

    /************************************************************
     * Download
     ************************************************************/
    download(workbook,meta){

        XLSX.writeFile(

            workbook,

            `Relatório de Eventos para o CPF ${meta.cpf}.xlsx`

        );

    }

}

/*****************************************************************
 * Inicialização
 *****************************************************************/

window.addEventListener("DOMContentLoaded",()=>{

    new Heidegger();

});