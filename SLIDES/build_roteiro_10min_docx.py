# /// script
# requires-python = ">=3.11"
# dependencies = [
#     "python-docx",
# ]
# ///

"""
Gerador do Roteiro de Apresentação de Iniciação Científica (10 Minutos)
Projeto: Do Barro ao Código: O Mito do Golem em Isaac Bashevis Singer e Helene Wecker
         sob a Paralaxe de Norbert Wiener (FFLCH-USP)
Gera:
  1) SLIDES/ROTEIRO_APRESENTACAO_10MIN_GOLEM.docx
  2) ROTEIRO_APRESENTACAO_10MIN_GOLEM.docx (cópia na raiz do projeto)
  3) SLIDES/ROTEIRO_APRESENTACAO_10MIN_GOLEM.md (espelho em Markdown)
"""

import os
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# ==============================================================================
# DADOS ESTRUTURADOS DO ROTEIRO (100% ALINHADOS AOS SLIDES 01-10 E AO PROJETO)
# ==============================================================================

META = {
    "instituicao": "UNIVERSIDADE DE SÃO PAULO (USP)",
    "unidade": "FACULDADE DE FILOSOFIA, LETRAS E CIÊNCIAS HUMANAS (FFLCH)",
    "departamento": "Departamento de Letras Orientais • Área de Estudos Judaicos e Árabes",
    "titulo_doc": "ROTEIRO DE APRESENTAÇÃO ORAL — INICIAÇÃO CIENTÍFICA (10 MINUTOS)",
    "titulo_pesquisa": "Do Barro ao Código: O Mito do Golem em Isaac Bashevis Singer e Helene Wecker sob a Paralaxe de Norbert Wiener",
    "subtitulo_pesquisa": "Estudo comparado de crítica textual, historicidade e representações do antissemitismo",
    "discente": "Eduardo Neri (Letras — Hebraico)",
    "orientador": "Prof. Dr. Luis Sérgio Krausz",
    "arquivo_slides": "index.html (10 slides interativos)",
}

AUTORES_E_OBRAS = [
    {
        "autor": "Norbert Wiener (1894–1964)",
        "papel_trialogo": "A Lente Metodológica (Paralaxe Cibernética)",
        "nota_grafia": "Nota de apoio oral: grafia correta nos slides e na bibliografia é Norbert Wiener (às vezes referido oralmente como \"Werner\"); pronuncia-se aproximadamente \"Uíner\".",
        "perfil": (
            "Matemático e filósofo norte-americano do MIT, de origem judaica (com tradição familiar "
            "associada à linhagem intelectual rabínica), fundador da Cibernética e pioneiro no debate ético "
            "sobre máquinas capazes de aprendizado e sistemas autônomos no pós-Segunda Guerra Mundial."
        ),
        "obra_titulo": "Deus & Golem, Cia. (God & Golem, Inc., 1964; ed. bras. Cultrix, 1968)",
        "obra_resumo": (
            "Ensaio filosófico curto e incisivo em que Wiener aproxima a mística cabalística da criação do "
            "Golem aos sistemas cibernéticos autorregulados. Sua tese central — usada na pesquisa como "
            "\"paralaxe\" analítica — demonstra que, ao conferir agência e capacidade de aprendizado a um "
            "artefato (seja de barro ou de código), o artífice perde inevitavelmente o monopólio de controle: "
            "a criatura deixa de obedecer cegamente e passa a produzir escolhas e consequências imprevisíveis."
        ),
    },
    {
        "autor": "Isaac Bashevis Singer (1902–1991)",
        "papel_trialogo": "A Raiz Asquenaze e a Tragédia do Criador",
        "nota_grafia": "",
        "perfil": (
            "Escritor judeu polonês radicado nos Estados Unidos a partir de 1935 e Prêmio Nobel de Literatura "
            "(1978). Filho e neto de rabinos, formado na mundividência chassídica da Europa Oriental, escreveu "
            "em língua iídiche sob o impacto do desaparecimento das comunidades asquenazes europeias, "
            "tensionando misticismo tradicional e dilemas morais modernos."
        ),
        "obra_titulo": "O Golem (The Golem, 1969/1982; ed. bras. Perspectiva, 1989)",
        "obra_resumo": (
            "Releitura lírica e trágica da lenda quinhentista do Maharal de Praga (Rabino Leib) sob o reinado "
            "de Rodolfo II. Para defender o gueto dos falsos libelos de sangue tramados pelo conde Bratislawski, "
            "o rabino anima um gigante da lama do rio Moldava inscrevendo Emet (Verdade) em sua testa. "
            "Cumprida a missão protetora, ocorre o colapso do plano teúrgico: o autômato desperta para a "
            "melancolia, descobre o afeto humano (Miriam/Elizabeth) e recusa retornar ao pó, expondo a culpa "
            "ética de um criador que despertou sensibilidade em um ser privado de fala e pertencimento."
        ),
    },
    {
        "autor": "Helene Wecker (1975–)",
        "papel_trialogo": "A Emancipação Diaspórica, Feminista e Intercultural",
        "nota_grafia": "",
        "perfil": (
            "Escritora judia norte-americana contemporânea, formada em escrita criativa pela Universidade "
            "Columbia. Em ensaios como \"The Most Jewish Thing I Do\", afirma uma identidade judaica secular "
            "fundada não na observância ortodoxa, mas no Midrash literário: contar histórias e reler criticamente "
            "as tradições para enfrentar os impasses da modernidade e da imigração."
        ),
        "obra_titulo": "A Golem e o Jinni / O Golem e o Gênio (The Golem and the Jinni, 2013; ed. bras. DarkSide, 2015)",
        "obra_resumo": (
            "Ambientado na Nova York de 1899, narra a jornada de Chava, uma Golem feminina moldada na Polônia "
            "por um cabalista heterodoxo para servir de esposa submissa a um comerciante. Com a morte súbita do "
            "mestre no navio transatlântico, ela aporta órfã de tutela no Lower East Side. Trabalhando em uma "
            "padaria e orientada pelo Rabino Meyer, Chava encontra Ahmad, um Jinni sírio feito de fogo e preso "
            "em corpo humano no enclave vizinho de Little Syria (Washington Street). Entre o barro (dever, peso) "
            "e o fogo (volúpia, liberdade), Chava rompe a servidão patriarcal e forja uma ética própria de solidariedade."
        ),
    },
]

MATRIZ_TRIALOGO = [
    ("Categoria do Triálogo", "Norbert Wiener (1964/1968)", "Isaac B. Singer (1969/1982)", "Helene Wecker (2013/2015)"),
    (
        "Função na Pesquisa",
        "Paralaxe crítica e epistemológica (o eixo do controle e da autonomia).",
        "Polo trágico-asquenaze (preservação da memória e limite do rabino).",
        "Polo diaspórico-polifônico (emancipação secular e alteridade).",
    ),
    (
        "Cenário Histórico",
        "Pós-guerra e corrida cibernética nos EUA (década de 1960).",
        "Gueto de Praga no séc. XVI (Imperador Rodolfo II), escrito no exílio.",
        "Nova York (1899): Lower East Side judaico e Little Syria árabe.",
    ),
    (
        "Natureza da Criatura",
        "Autômato / máquina capaz de aprendizado e resposta ao ambiente.",
        "Gigante masculino de barro do rio Moldava; mudo, forte e melancólico.",
        "Golem feminina (Chava); trabalhadora urbana, empática e reflexiva.",
    ),
    (
        "A Crise de Tutela",
        "\"O servo que aprende destrói a obediência cega ao programador.\"",
        "O Rabino Leib perde o controle moral quando o barro sente dor e amor.",
        "O mestre morre no mar; Chava transforma o Emet de submissão em arbítrio.",
    ),
    (
        "Resposta ao Preconceito",
        "Alerta ético contra delegar justiça e violência a autômatos cegos.",
        "Autodefesa física e jurídica contra libelos de sangue antissemitas.",
        "Convívio intercultural com o Jinni sírio diante do nativismo americano.",
    ),
]

SLIDES_ROTEIRO = [
    {
        "num": "01",
        "titulo": "Do Barro ao Código: A Ontologia do Autômato",
        "eixo": "Eixo 01 • Epistemologia e Ontologia",
        "tempo_duracao": "60s",
        "cronometro": "00:00 – 01:00",
        "lider": "Norbert Wiener (com base cabalística de Moshe Idel)",
        "tela_resumo": (
            "Imagem: a faísca da agência entre matéria e código. "
            "Citação central: WIENER (1968, p. 45) sobre a perda de controle ao conferir agência. "
            "Cards inferiores: (1) Teurgia Talmúdica (Sefer Yetzirah / IDEL, 1990); (2) Paralaxe Cibernética."
        ),
        "acao_tela": "Iniciar em Tela Cheia (tecla F). Manter o olhar na banca ao abrir e apontar brevemente o conceito de 'Paralaxe' na tela.",
        "bullets": [
            "Saudação rápida + apresentação do título e do triálogo (Singer, Wecker e Wiener).",
            "Explicar a matriz mística (Sefer Yetzirah, permutação de letras, Moshe Idel).",
            "Introduzir Wiener como lente de 'paralaxe': quando se dá agência ao artefato, perde-se a tutela.",
        ],
        "fala": (
            "Bom dia a todos. Sou **Eduardo Neri**, de Letras-Hebraico na FFLCH-USP, e apresento minha Iniciação "
            "Científica orientada pelo **Prof. Dr. Luis Sérgio Krausz**: \"Do Barro ao Código: O Mito do Golem em "
            "Isaac Bashevis Singer e Helene Wecker sob a Paralaxe de Norbert Wiener\". Nossa pesquisa parte de uma "
            "questão: o que ocorre quando uma criação ganha agência e escapa ao seu criador? Na mística medieval "
            "estudada por **Moshe Idel**, animar o barro pelas letras do *Sefer Yetzirah* era um ato teúrgico. Mas, "
            "ao ler a ficção moderna em um **triálogo**, adotamos **Norbert Wiener**, pai da cibernética, como lente "
            "de **paralaxe**. Como diz a citação na tela, conferir agência a um artefato implica a **perda de controle "
            "do artífice** — do barro ao código, a criatura rompe a obediência cega. "
            "*Mas como esse rito místico se tornou resposta histórica à violência? Vejamos em Praga.*"
        ),
    },
    {
        "num": "02",
        "titulo": "O Gueto e a Lama: A Gênese Asquenaze",
        "eixo": "Eixo 02 • Historicidade e Autodefesa",
        "tempo_duracao": "60s",
        "cronometro": "01:00 – 02:00",
        "lider": "Isaac Bashevis Singer + Historicidade (Maharal de Praga / Moshe Idel)",
        "tela_resumo": (
            "Imagem: viela opressiva de Praga com figura colossal de barro. "
            "Citação central: Praga histórica do imperador Rodolfo II e o Maharal (SINGER, 1989; IDEL, 1990). "
            "Cards inferiores: (1) Pano de Fundo Real (Rodolfo II); (2) O Mito como Resposta (autodefesa biopolítica)."
        ),
        "acao_tela": "Pressionar → para o Slide 02. Destacar visualmente a atmosfera sombria do gueto de Praga.",
        "bullets": [
            "Situar a Praga do século XVI sob o imperador Rodolfo II (ambiente hermético e alquímico real).",
            "Explicar o papel do Maharal (Rabino Judá Loew) e a ameaça dos libelos de sangue.",
            "Mostrar como Singer ancora o mito na sobrevivência física da comunidade asquenaze.",
        ],
        "fala": (
            "Na tradição asquenaze da Europa Central, o Golem deixa de ser exercício contemplativo e vira muralha "
            "física. Estamos na **Praga do século XVI**, sob o imperador **Rodolfo II**, corte real voltada à alquimia "
            "e ao hermetismo onde o **Rabino Judá Loew** — o **Maharal** — foi de fato recebido pelo monarca. Ao "
            "retomar essa memória séculos depois em *O Golem*, **Isaac Bashevis Singer** ancora a lenda na carne viva "
            "do gueto. A lama do rio Moldava não é erguida por soberba, mas por uma urgência extrema: a **autodefesa "
            "física e jurídica** contra os **libelos de sangue**, falsas acusações de crime ritual que geravam "
            "perseguições antissemitas. Em Singer, o Golem nasce como um escudo telúrico e mudo dos oprimidos. "
            "*Contudo, no romance contemporâneo de Helene Wecker, esse mito fechado no gueto passa por uma virada radical.*"
        ),
    },
    {
        "num": "03",
        "titulo": "Choque Ontológico: O Barro e o Fogo",
        "eixo": "Eixo 03 • Literatura Comparada",
        "tempo_duracao": "60s",
        "cronometro": "02:00 – 03:00",
        "lider": "Helene Wecker + Literatura Comparada (Tânia Franco Carvalhal)",
        "tela_resumo": (
            "Imagem: figura feminina de argila caminhando ao lado de silhueta de chamas. "
            "Citação central: fricção entre duas cosmologias orientais no exílio (CARVALHAL, 1986; WECKER, 2015). "
            "Cards inferiores: (1) A Matéria da Servidão (Barro/Chava); (2) A Chama da Liberdade (Fogo/Ahmad)."
        ),
        "acao_tela": "Pressionar → para o Slide 03. Fazer o gesto de contraste entre 'terra/dever' e 'fogo/liberdade'.",
        "bullets": [
            "Apresentar Helene Wecker e o romance *A Golem e o Jinni* (2013/2015).",
            "Contrastar as duas ontologias: Chava (barro = dever, contenção) vs. Ahmad (fogo = volúpia, liberdade).",
            "Explicar o método da Literatura Comparada (Tânia Carvalhal) no cotejo entre os dois folclores.",
        ],
        "fala": (
            "Entra aqui o segundo vértice ficcional do triálogo: **Helene Wecker**, com *A Golem e o Jinni*. Pelo "
            "método da **Literatura Comparada** de **Tânia Franco Carvalhal**, vemos como Wecker coloca em fricção "
            "duas cosmologias orientais deslocadas para o exílio. De um lado está **Chava**, uma Golem feminina "
            "moldada do **barro**: densa, prudente, presa à materialidade da terra e ávida por mandamentos que lhe "
            "justifiquem a existência. Do outro está **Ahmad**, um **Jinni** sírio forjado do **fogo** sem fumaça: "
            "volátil, apaixonado e avesso a qualquer domesticidade, aprisionado contra a vontade em um corpo humano. "
            "O choque ontológico entre o barro e o fogo encena a tensão entre o dever protetor das raízes e o desejo "
            "indomável de liberdade individual no mundo secular. "
            "*E onde essas forças se encontram? Numa cartografia histórica muito precisa.*"
        ),
    },
    {
        "num": "04",
        "titulo": "Diáspora e o Novo Mundo: Cartografia de 1899",
        "eixo": "Eixo 04 • Cartografia, Diáspora e Espaço",
        "tempo_duracao": "70s",
        "cronometro": "03:00 – 04:10",
        "lider": "Helene Wecker + Cartografia Histórica (Linda K. Jacobs / David Rumsey 1899)",
        "tela_resumo": (
            "Imagem inicial: Mapa histórico David Rumsey (1899) com rotas transatlânticas; alternável para cenário urbano de Manhattan. "
            "Citação central: Lower East Side e o enclave real de Little Syria (JACOBS, 2015; WECKER, 2015). "
            "Cards inferiores: (1) Little Syria Real (Washington St.); (2) A Ruptura Transatlântica."
        ),
        "acao_tela": "Pressionar → para o Slide 04. Enquanto fala das rotas marítimas, manter o mapa; ao citar o Lower East Side e Washington Street, clicar no botão '🏙️ Cenário Urbano (Manhattan 1899)' (e opcionalmente 1 clique rápido na Lupa se o tempo permitir).",
        "bullets": [
            "Mostrar a base documental: mapas de 1899 da coleção David Rumsey (rotas Praga/Hamburgo e Levante).",
            "Morte do mestre no navio: Chava chega ao Lower East Side sem linhagem ou tutela rabínica.",
            "Citar Linda K. Jacobs: Little Syria em Washington Street era um bairro vizinho real ao gueto judaico.",
        ],
        "fala": (
            "Esse deslocamento tem lastro documental. A partir da cartografia histórica de 1899 da **David Rumsey "
            "Collection**, observamos as rotas dos vapores transatlânticos que ligavam os portos da Europa Central "
            "e do Levante otomano a Nova York. No romance, a travessia oceânica dissolve a autoridade do Velho Mundo: "
            "o cabalista heterodoxo Yehudah Schaalman morre subitamente a bordo do navio, e Chava aporta no **Lower "
            "East Side** completamente sozinha e sem senhor. [CLICAR NO BOTÃO 'Cenário Urbano (Manhattan 1899)'] "
            "Como documenta a historiadora **Linda K. Jacobs**, a poucos minutos de caminhada dos cortiços judaicos "
            "florescia então a colônia árabe real de **Little Syria**, sediada em **Washington Street**. Sem a "
            "ameaça dos pogroms europeus, o novo \"gueto\" torna-se a metrópole imigrante — um laboratório secular "
            "onde a criatura perde a tutela rabínica e precisa sobreviver como trabalhadora urbana. "
            "*É exatamente nesse vácuo de autoridade que a voz de Norbert Wiener se torna decisiva.*"
        ),
    },
    {
        "num": "05",
        "titulo": "Deus & Golem, Cia.: A Perda do Monopólio",
        "eixo": "Eixo 05 • Teoria Cibernética",
        "tempo_duracao": "65s",
        "cronometro": "04:10 – 05:15",
        "lider": "Norbert Wiener (eixo articulador do triálogo)",
        "tela_resumo": (
            "Imagem: engrenagens cibernéticas envolvendo um pergaminho hebraico. "
            "Citação central: WIENER (1968, p. 82) sobre criar uma máquina à nossa imagem. "
            "Cards inferiores: (1) O Paradoxo do Servo; (2) A Lição Humanista."
        ),
        "acao_tela": "Pressionar → para o Slide 05. Pausar meio segundo na leitura da citação de Wiener.",
        "bullets": [
            "Apresentar o livro *Deus & Golem, Cia.* (1964; trad. bras. 1968) como centro do triálogo.",
            "Explicar o 'Paradoxo do Servo': eficiência exige aprendizado/autonomia; autonomia destrói a servidão.",
            "Conectar a advertência de Wiener ao risco moral tanto da tecnologia quanto dos guardiões literários.",
        ],
        "fala": (
            "Chegamos ao eixo teórico do nosso triálogo. Em seu ensaio *Deus & Golem, Cia.*, de 1964, o matemático "
            "**Norbert Wiener** recorre diretamente ao mito cabalístico para lançar um alerta ético à era da "
            "automação. Como destaca a epígrafe em tela: se criarmos máquinas capazes de aprender à nossa imagem, "
            "elas não serão menos perigosas para nós do que nós próprios fomos para Deus. Wiener expõe aquilo que "
            "chamamos de **Paradoxo do Servo**: para agir com eficácia em um ambiente complexo e imprevisível, o "
            "artefato precisa receber autonomia e aprendizado por *feedback*; porém, essa exata liberdade destrói a "
            "submissão cega desejada pelo programador ou rabino. Delegar a proteção humana a autômatos sem piedade "
            "expõe a ilusão da técnica. *Vejamos como Singer e Wecker encenam essa crise de controle em direções "
            "opostas — começando por Singer.*"
        ),
    },
    {
        "num": "06",
        "titulo": "A Melancolia do Ser: Isaac Bashevis Singer",
        "eixo": "Eixo 06 • Crítica Textual e Tragédia",
        "tempo_duracao": "70s",
        "cronometro": "05:15 – 06:25",
        "lider": "Isaac Bashevis Singer (sob a paralaxe de Wiener)",
        "tela_resumo": (
            "Imagem: close-up no rosto de barro rachando e revelando um olho triste. "
            "Citação central: SINGER (1989, p. 64) sobre o Golem olhando o céu e sentindo o frio da solidão. "
            "Cards inferiores: (1) A Falência do Rabino; (2) A Inadequação Existencial."
        ),
        "acao_tela": "Pressionar → para o Slide 06. Tom de voz mais introspectivo ao falar da melancolia do Golem.",
        "bullets": [
            "Descrever o *close reading* da novela de Singer: a criatura não é um monstro violento, mas melancólico.",
            "Mostrar o conflito do Rabino Leib: falência ética ao dar sensibilidade sem conceder fala nem alma.",
            "Alegoria da condição judaica na Europa: útil na hora do perigo, temido e descartado na paz.",
        ],
        "fala": (
            "Sob a paralaxe de Wiener, a prosa em iídiche de **Isaac Bashevis Singer** revela uma perda de tutela "
            "profundamente **trágica**. Escrevendo após a destruição do judaísmo europeu, Singer mostra o Golem "
            "cumprindo sua missão em Praga contra o conde Bratislawski, mas algo escapa ao cálculo teúrgico: pelas "
            "rachaduras do barro desperta uma **consciência melancólica**. Ele contempla o céu, escuta os pássaros, "
            "sente o frio da solidão, descobre o afeto humano e recusa retornar docilmente ao pó. Isso expõe a "
            "**falência ética do Rabino Leib**, atormentado pelo remorso de ter arrancado a matéria do repouso, "
            "concedendo-lhe dor e desejo sem lhe dar fala articulada nem alma plena. O gigante desajeitado torna-se "
            "emblema da diáspora europeia: necessário na hora da crise, mas temido e segregado assim que a paz retorna. "
            "*Se em Singer a autonomia termina em luto, em Helene Wecker ela se converte em emancipação.*"
        ),
    },
    {
        "num": "07",
        "titulo": "A Emancipação do Barro: Helene Wecker",
        "eixo": "Eixo 07 • Subversão Feminista",
        "tempo_duracao": "70s",
        "cronometro": "06:25 – 07:35",
        "lider": "Helene Wecker (em contraponto a Singer e Wiener)",
        "tela_resumo": (
            "Imagem: mulher apagando os fios invisíveis e ressignificando o 'Emet'. "
            "Citação central: esposa perfeita de barro que descobre o poder de dizer 'não' (WECKER, 2015). "
            "Cards inferiores: (1) A Golem como Mulher; (2) A Verdade Redefinida (Emet ético)."
        ),
        "acao_tela": "Pressionar → para o Slide 07. Enfatizar a ruptura de gênero e o termo hebraico 'Emet'.",
        "bullets": [
            "Subversão do cânone androcêntrico: a Golem torna-se mulher (Chava), originalmente encomendada como esposa submissa.",
            "Com a morte do mestre, Chava controla sua força, torna-se padeira e aprende a dizer 'não'.",
            "Ressignificação do *Emet* (Verdade): deixa de ser obediência ao mestre e vira responsabilidade moral.",
        ],
        "fala": (
            "Se Singer explora o remorso do criador, **Helene Wecker** narra o triunfo ético da criatura. Sua "
            "protagonista, **Chava**, subverte séculos de tradição rabínica androcêntrica: ela não nasce para "
            "defender um gueto, mas é encomendada por um comerciante que desejava uma esposa autômata, silenciosa "
            "e incapaz de contestar. Com a morte do mestre no mar, esse pacto patriarcal colapsa antes de se "
            "consumar, e Chava descobre em Nova York o poder vertiginoso de **dizer \"não\"**. Dotada de uma empatia "
            "extrema que capta os anseios alheios, ela aprende a domar a própria força física, sustenta-se como "
            "padeira sob a escuta ética do Rabino Meyer e assume o próprio destino. A palavra cabalística **Emet** — "
            "Verdade — que antes subordinava o Golem ao mestre, transmuta-se aqui em **autenticidade interior** e "
            "livre-arbítrio da mulher imigrante. *E é essa autonomia moral que viabiliza o diálogo com o outro marginalizado.*"
        ),
    },
    {
        "num": "08",
        "titulo": "O Encontro na Alteridade: Little Syria",
        "eixo": "Eixo 08 • Convivência e Alteridade",
        "tempo_duracao": "65s",
        "cronometro": "07:35 – 08:40",
        "lider": "Helene Wecker (em diálogo com Carvalhal e Jacobs)",
        "tela_resumo": (
            "Imagem: chá árabe de hortelã e café nova-iorquino sobre manuscritos. "
            "Citação central: fatura polifônica articulando gueto judaico e Little Syria (PROJETO IC; JACOBS, 2015; CARVALHAL, 1986). "
            "Cards inferiores: (1) Polifonia Cultural; (2) O Monstro como Espelho."
        ),
        "acao_tela": "Pressionar → para o Slide 08. Destacar a transição da autodefesa fechada para o convívio intercultural.",
        "bullets": [
            "Passagem da resposta física ao preconceito (Praga) para a solidariedade interétnica (Nova York).",
            "Caminhadas noturnas de Chava e Ahmad: judeus e árabes unidos pelo desabrigo diante do nativismo.",
            "Polifonia literária: saberes iídiches e tradições levantinas coexistem sem apagar as diferenças.",
        ],
        "fala": (
            "Nisso reside a resposta original de Wecker ao preconceito e à xenofobia. Enquanto em Praga o "
            "antissemitismo exigia uma muralha de barro voltada para dentro do gueto, na Nova York de 1899 a "
            "narrativa aposta na **polifonia** e no encontro com a **alteridade**. Chava, que não dorme, e o Jinni "
            "sírio Ahmad, igualmente insone, percorrem juntos as ruas noturnas de Manhattan entre o Lower East Side "
            "e **Little Syria**. Diante de um nativismo americano que enxergava tanto judeus asquenazes quanto árabes "
            "levantinos como corpos estranhos e \"monstros\", as duas criaturas reconhecem-se na mesma solidão de "
            "exilados. A mesa onde coexistem o chá de hortelã sírio e o café nova-iorquino não apaga suas tradições "
            "distintas, mas funda uma **ética da convivência**, substituindo a retaliação pela solidariedade entre "
            "minorias. *Assim chegamos à síntese do que permanece quando a magia cessa.*"
        ),
    },
    {
        "num": "09",
        "titulo": "O Texto como Resistência: A Força da Literatura",
        "eixo": "Eixo 09 • Síntese Hermenêutica",
        "tempo_duracao": "55s",
        "cronometro": "08:40 – 09:35",
        "lider": "Síntese do Triálogo (Singer + Wecker + Wiener + Auerbach/Carvalhal)",
        "tela_resumo": (
            "Imagem: livro antigo aberto cujas letras iluminam uma cidade moderna. "
            "Citação central: maleabilidade do mito para absorver as feridas da história (CARVALHAL, 1986; AUERBACH, 1971). "
            "Cards inferiores: (1) Sobrevivência Simbólica; (2) O Papel da Crítica Textual."
        ),
        "acao_tela": "Pressionar → para o Slide 09. Ritmo firme e conclusivo, amarrando os três autores juntos.",
        "bullets": [
            "Amarrar o triálogo: Wiener revela a crise da agência; Singer, o luto trágico; Wecker, a autonomia secular.",
            "Auerbach e Carvalhal: o mito sobrevive porque muda de forma para elaborar feridas históricas novas.",
            "A literatura substitui a argila: o texto escrito é o verdadeiro rito contemporâneo de preservação da memória.",
        ],
        "fala": (
            "Concluindo, o triálogo entre **Singer**, **Wecker** e **Norbert Wiener** demonstra por que esse mito "
            "atravessa os séculos: do sótão da sinagoga de Praga aos autômatos pensantes do século XX e às ruas "
            "imigrantes de Manhattan. Em Singer, a perda de controle revela o luto e a culpa moral do criador; em "
            "Wecker, a mesma ruptura abre caminho para a emancipação feminina e para o diálogo intercultural. Como "
            "mostram **Erich Auerbach** e **Tânia Carvalhal**, a mimesis literária garante a sobrevida do mito porque "
            "sua forma é maleável o suficiente para absorver as feridas históricas de cada geração. Quando a argila se "
            "esfarela e os ritos silenciam, é o próprio **código literário** que preserva a memória judaica e a "
            "converte em espelho universal dos dilemas da autonomia humana."
        ),
    },
    {
        "num": "10",
        "titulo": "Aparato Crítico: Referências Bibliográficas",
        "eixo": "Eixo 10 • Fundamentação Científica & Créditos",
        "tempo_duracao": "25s",
        "cronometro": "09:35 – 10:00",
        "lider": "Corpus Primário e Fortuna Crítica Completa",
        "tela_resumo": (
            "Layout em duas colunas: Corpus Literário Primário (Singer, Wecker), Cibernética (Wiener), "
            "Mística/Cartografia/História (Idel, Jacobs, David Rumsey) e Literatura Comparada (Carvalhal, Auerbach)."
        ),
        "acao_tela": "Pressionar → para o Slide 10. Manter projetado durante o encerramento e eventual arguição da banca.",
        "bullets": [
            "Apontar rapidamente a divisão do aparato: Corpus Primário + Cibernética + História/Cartografia + Teoria Literária.",
            "Agradecer à banca, à FFLCH-USP e ao orientador Prof. Dr. Luis Sérgio Krausz.",
            "Abrir espaço para perguntas.",
        ],
        "fala": (
            "Encerro apresentando o aparato crítico da pesquisa: o corpus primário de **Singer** e **Wecker**, a "
            "cibernética de **Wiener**, a historiografia e cartografia de **Idel**, **Jacobs** e **David Rumsey**, e "
            "a teoria literária de **Carvalhal** e **Auerbach**. Muito obrigado à banca, à **FFLCH-USP** e ao meu "
            "orientador, **Prof. Dr. Luis Sérgio Krausz**."
        ),
    },
]

PERGUNTAS_BANCA = [
    (
        "1. Por que usar um matemático ciberneticista (Norbert Wiener) em uma pesquisa de Letras e Estudos Judaicos?",
        "Porque Wiener não trata o Golem como simples metáfora decorativa: em 'Deus & Golem, Cia.' (1964), ele analisa o mito a partir da própria tradição judaica (sendo inclusive descendente de linhagem intelectual ligada a Maimônides, como ele mesmo lembra) para discutir a ética da criação autônoma. Como recurso de 'paralaxe', Wiener oferece uma segunda linha de visada que ilumina nos textos de Singer e Wecker exatamente o ponto em que a criatura deixa de ser extensão do mestre e se torna sujeito moral imprevisível."
    ),
    (
        "2. O que muda na representação do antissemitismo e do preconceito entre Singer e Wecker?",
        "Em Singer (ancorado na Praga quinhentista de Rodolfo II e escrito após a Shoá), o antissemitismo aparece sob a forma direta da violência física e dos libelos de sangue, exigindo um guardião de barro dentro dos muros do gueto — mas cuja existência revela o trauma e o isolamento. Em Wecker (Nova York, 1899), o preconceito assume a feição da xenofobia urbana e da pressão assimilacionista contra imigrantes judeus e árabes; a resposta deixa de ser a retaliação defensiva e passa a ser a construção ética da convivência entre o Lower East Side e Little Syria."
    ),
    (
        "3. Como a materialidade da língua e da escrita aparece nessa passagem do barro para o texto?",
        "Na matriz do Sefer Yetzirah (Idel), a vida vem da permutação de letras hebraicas e da palavra 'Emet' na fronte. Em Singer, escrito originalmente em iídiche, a tragédia reside no fato de a criatura sentir afetos humanos sem possuir a linguagem articulada para expressá-los. Em Wecker, a própria autoria assume o judaísmo secular como prática de leitura e escrita (Midrash moderno), e Chava rompe o código da servidão gravado em sua gênese para reescrever a própria conduta no mundo."
    ),
]


# ==============================================================================
# UTILITÁRIOS DE FORMATAÇÃO PYTHON-DOCX
# ==============================================================================

COLOR_NAVY = RGBColor(15, 23, 42)        # #0F172A - Títulos e cabeçalhos
COLOR_GOLD_DARK = RGBColor(146, 102, 22) # #926616 - Subtítulos e acentos
COLOR_SLATE = RGBColor(71, 85, 105)      # #475569 - Metadados e notas
COLOR_DARK_TEXT = RGBColor(30, 41, 59)   # #1E293B - Corpo de texto
COLOR_WHITE = RGBColor(255, 255, 255)

HEX_NAVY = "0F172A"
HEX_GOLD_BG = "FBF7EE"
HEX_SLATE_BG = "F1F5F9"
HEX_BORDER_GOLD = "C59B38"
HEX_BORDER_LIGHT = "CBD5E1"
HEX_ZEBRA = "F8FAFC"


def set_cell_background(cell, hex_color: str):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_borders(table, color=HEX_BORDER_LIGHT, sz="4"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f"</w:tblBorders>"
    )
    tblPr.append(borders)


def add_left_border_box(doc, title: str, text_with_bold: str, bg_hex: str, border_hex: str, space_after=8):
    """Cria uma caixa destacada (tabela de 1 célula) com borda esquerda grossa para falas ou avisos."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Cm(17.0)
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=180)

    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER_LIGHT}"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER_LIGHT}"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER_LIGHT}"/>'
        f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{border_hex}"/>'
        f"</w:tcBorders>"
    )
    tcPr.append(borders)

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.25
    r_title = p.add_run(title)
    r_title.bold = True
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(10)
    r_title.font.color.rgb = COLOR_GOLD_DARK

    p_body = cell.add_paragraph()
    p_body.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_body.paragraph_format.space_before = Pt(0)
    p_body.paragraph_format.space_after = Pt(0)
    p_body.paragraph_format.line_spacing = 1.38
    add_formatted_runs(p_body, text_with_bold, base_size=11.5, base_color=COLOR_DARK_TEXT)

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(space_after)


def add_formatted_runs(paragraph, text: str, base_size=11, base_color=COLOR_DARK_TEXT):
    """Interpreta **negrito** e *itálico* simples dentro de uma string para runs do python-docx."""
    # Quebra mantendo marcadores **...** e *...*
    tokens = re.split(r"(\*\*.*?\*\*|\*[^*]+?\*)", text)
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**"):
            run = paragraph.add_run(tok[2:-2])
            run.bold = True
            run.font.color.rgb = COLOR_NAVY
        elif tok.startswith("*") and tok.endswith("*"):
            run = paragraph.add_run(tok[1:-1])
            run.italic = True
            run.font.color.rgb = COLOR_GOLD_DARK if tok.startswith("*Mas") or tok.startswith("*E ") or tok.startswith("*Se ") or tok.startswith("*Isso") or tok.startswith("*Contudo") else base_color
        else:
            run = paragraph.add_run(tok)
            run.font.color.rgb = base_color
        run.font.name = "Calibri"
        run.font.size = Pt(base_size)


def add_section_heading(doc, text: str, page_break_before=False):
    if page_break_before:
        doc.add_page_break()
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True

    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="12" w:space="4" w:color="{HEX_BORDER_GOLD}"/>'
        f"</w:pBdr>"
    )
    pPr.append(pBdr)

    run = p.add_run(text)
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(14)
    run.font.color.rgb = COLOR_NAVY
    return p


def add_page_number_to_run(run):
    fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls("w"))
    instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls("w"))
    fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls("w"))
    fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls("w"))
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)


def count_spoken_words() -> int:
    total = 0
    for s in SLIDES_ROTEIRO:
        clean = re.sub(r"\[.*?\]", "", s["fala"])
        clean = clean.replace("**", "").replace("*", "")
        total += len(clean.split())
    return total


# ==============================================================================
# CONSTRUTOR DO DOCUMENTO DOCX
# ==============================================================================

def build_docx(output_paths: list[str]):
    doc = Document()

    # Configuração de página A4 com margens confortáveis para bancada
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(1.8)
        section.bottom_margin = Cm(1.8)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)

        # Cabeçalho discreto (não na primeira página)
        section.different_first_page_header_footer = True
        hdr = section.header
        p_hdr = hdr.paragraphs[0]
        p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_hdr = p_hdr.add_run("FFLCH-USP • Roteiro 10 min | Do Barro ao Código (Singer • Wecker • Wiener)")
        r_hdr.font.name = "Calibri"
        r_hdr.font.size = Pt(8.5)
        r_hdr.font.color.rgb = COLOR_SLATE

        # Rodapé com paginação
        ftr = section.footer
        p_ftr = ftr.paragraphs[0]
        p_ftr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_f1 = p_ftr.add_run("Eduardo Neri • Iniciação Científica (Estudos Judaicos e Árabes)   |   Pág. ")
        r_f1.font.name = "Calibri"
        r_f1.font.size = Pt(8.5)
        r_f1.font.color.rgb = COLOR_SLATE
        r_pg = p_ftr.add_run()
        r_pg.font.name = "Calibri"
        r_pg.font.size = Pt(8.5)
        r_pg.font.color.rgb = COLOR_SLATE
        add_page_number_to_run(r_pg)

    # Estilo Normal padrão
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.font.color.rgb = COLOR_DARK_TEXT

    # --------------------------------------------------------------------------
    # CABEÇALHO INSTITUCIONAL (BANNER DE ABERTURA)
    # --------------------------------------------------------------------------
    banner = doc.add_table(rows=1, cols=1)
    banner.alignment = WD_TABLE_ALIGNMENT.CENTER
    banner.autofit = False
    b_cell = banner.cell(0, 0)
    b_cell.width = Cm(17.0)
    set_cell_background(b_cell, HEX_NAVY)
    set_cell_margins(b_cell, top=180, bottom=180, left=220, right=220)

    p1 = b_cell.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_after = Pt(2)
    r1 = p1.add_run(f"{META['instituicao']}  •  {META['unidade']}\n{META['departamento']}")
    r1.font.name = "Calibri"
    r1.font.size = Pt(9.5)
    r1.font.color.rgb = RGBColor(226, 183, 104)
    r1.bold = True

    p2 = b_cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(6)
    p2.paragraph_format.space_after = Pt(4)
    r2 = p2.add_run(META["titulo_doc"])
    r2.font.name = "Calibri"
    r2.font.size = Pt(13)
    r2.font.color.rgb = COLOR_WHITE
    r2.bold = True

    p3 = b_cell.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_after = Pt(6)
    r3 = p3.add_run(META["titulo_pesquisa"])
    r3.font.name = "Calibri"
    r3.font.size = Pt(11.5)
    r3.font.color.rgb = RGBColor(241, 245, 249)
    r3.italic = True

    p4 = b_cell.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_after = Pt(0)
    r4 = p4.add_run(f"Discente: {META['discente']}   |   Orientador: {META['orientador']}")
    r4.font.name = "Calibri"
    r4.font.size = Pt(9.5)
    r4.font.color.rgb = RGBColor(203, 213, 225)

    total_words = count_spoken_words()
    p_meta_tip = doc.add_paragraph()
    p_meta_tip.paragraph_format.space_before = Pt(8)
    p_meta_tip.paragraph_format.space_after = Pt(6)
    add_formatted_runs(
        p_meta_tip,
        f"**Como usar este roteiro na apresentação:** Os blocos de fala somam **{total_words} palavras** "
        f"(ritmo natural de **~130 a 140 palavras/minuto**, fechando exatamente em **10 minutos** com pausas de transição). "
        f"Os termos em **negrito** servem de faróis visuais caso você desvie o olhar para os slides. "
        f"No Slide 04, há indicação exata do momento de alternar o mapa da *David Rumsey Collection* para a vista urbana de Manhattan.",
        base_size=10,
        base_color=COLOR_SLATE,
    )

    # --------------------------------------------------------------------------
    # SEÇÃO 1: O TRIÁLOGO — AUTORES E BREVÍSSIMO RESUMO DAS OBRAS
    # --------------------------------------------------------------------------
    add_section_heading(doc, "1. O TRIÁLOGO DA PESQUISA: AUTORES E BREVÍSSIMO RESUMO DAS OBRAS")

    for item in AUTORES_E_OBRAS:
        p_a = doc.add_paragraph()
        p_a.paragraph_format.space_before = Pt(6)
        p_a.paragraph_format.space_after = Pt(2)
        p_a.paragraph_format.keep_with_next = True
        ra1 = p_a.add_run(item["autor"])
        ra1.bold = True
        ra1.font.name = "Calibri"
        ra1.font.size = Pt(12)
        ra1.font.color.rgb = COLOR_NAVY
        ra2 = p_a.add_run(f"  —  {item['papel_trialogo']}")
        ra2.bold = True
        ra2.font.name = "Calibri"
        ra2.font.size = Pt(10.5)
        ra2.font.color.rgb = COLOR_GOLD_DARK

        if item["nota_grafia"]:
            p_n = doc.add_paragraph()
            p_n.paragraph_format.space_before = Pt(0)
            p_n.paragraph_format.space_after = Pt(2)
            rn = p_n.add_run(item["nota_grafia"])
            rn.italic = True
            rn.font.name = "Calibri"
            rn.font.size = Pt(9.5)
            rn.font.color.rgb = COLOR_GOLD_DARK

        p_perf = doc.add_paragraph()
        p_perf.paragraph_format.space_before = Pt(0)
        p_perf.paragraph_format.space_after = Pt(2)
        p_perf.paragraph_format.line_spacing = 1.2
        add_formatted_runs(p_perf, f"**Perfil:** {item['perfil']}", base_size=10.5)

        p_obr = doc.add_paragraph()
        p_obr.paragraph_format.space_before = Pt(0)
        p_obr.paragraph_format.space_after = Pt(6)
        p_obr.paragraph_format.line_spacing = 1.2
        add_formatted_runs(p_obr, f"**Obra ({item['obra_titulo']}):** {item['obra_resumo']}", base_size=10.5)

    # Tabela-Síntese do Triálogo
    p_mat_title = doc.add_paragraph()
    p_mat_title.paragraph_format.space_before = Pt(6)
    p_mat_title.paragraph_format.space_after = Pt(4)
    p_mat_title.paragraph_format.keep_with_next = True
    rm = p_mat_title.add_run("Síntese Comparativa do Triálogo (Como os Três Autores Conversam nos Slides)")
    rm.bold = True
    rm.font.name = "Calibri"
    rm.font.size = Pt(11)
    rm.font.color.rgb = COLOR_NAVY

    tbl_mat = doc.add_table(rows=len(MATRIZ_TRIALOGO), cols=4)
    tbl_mat.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_mat.autofit = False
    set_table_borders(tbl_mat, HEX_BORDER_LIGHT, "4")
    col_widths_mat = [Cm(3.2), Cm(4.6), Cm(4.6), Cm(4.6)]

    for r_idx, row_data in enumerate(MATRIZ_TRIALOGO):
        row = tbl_mat.rows[r_idx]
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = col_widths_mat[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(val)
            run.font.name = "Calibri"
            if r_idx == 0:
                set_cell_background(cell, HEX_NAVY)
                run.bold = True
                run.font.size = Pt(9.5)
                run.font.color.rgb = COLOR_WHITE
            else:
                if r_idx % 2 == 0:
                    set_cell_background(cell, HEX_ZEBRA)
                if c_idx == 0:
                    run.bold = True
                    run.font.color.rgb = COLOR_NAVY
                else:
                    run.font.color.rgb = COLOR_DARK_TEXT
                run.font.size = Pt(9.2)

    # --------------------------------------------------------------------------
    # SEÇÃO 2: MAPA DE VOO — CRONOMETRAGEM DOS 10 MINUTOS
    # --------------------------------------------------------------------------
    add_section_heading(doc, "2. MAPA DE VOO E GESTÃO DO TEMPO (00:00 – 10:00)", page_break_before=True)

    p_map_intro = doc.add_paragraph()
    p_map_intro.paragraph_format.space_after = Pt(6)
    add_formatted_runs(
        p_map_intro,
        "Acompanhe o relógio pelos marcos acumulados abaixo. **Meta de ritmo:** mantenha cerca de 1 minuto por slide "
        "na introdução (Slides 1–3), ganhe 10 segundos extras nos slides centrais de análise e cartografia (Slides 4, 6 e 7) "
        "e feche as referências em 25 segundos.",
        base_size=10,
        base_color=COLOR_SLATE,
    )

    tbl_time = doc.add_table(rows=len(SLIDES_ROTEIRO) + 1, cols=5)
    tbl_time.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_time.autofit = False
    set_table_borders(tbl_time, HEX_BORDER_LIGHT, "4")
    col_widths_time = [Cm(1.4), Cm(2.4), Cm(1.4), Cm(6.3), Cm(5.5)]

    headers_time = ["Slide", "Cronômetro", "Duração", "Título do Slide", "Foco no Triálogo"]
    for c_idx, h in enumerate(headers_time):
        cell = tbl_time.rows[0].cells[c_idx]
        cell.width = col_widths_time[c_idx]
        set_cell_background(cell, HEX_NAVY)
        set_cell_margins(cell, top=80, bottom=80, left=80, right=80)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.bold = True
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.font.color.rgb = COLOR_WHITE

    for idx, s in enumerate(SLIDES_ROTEIRO, start=1):
        row = tbl_time.rows[idx]
        vals = [s["num"], s["cronometro"], s["tempo_duracao"], s["titulo"], s["lider"]]
        for c_idx, v in enumerate(vals):
            cell = row.cells[c_idx]
            cell.width = col_widths_time[c_idx]
            if idx % 2 == 0:
                set_cell_background(cell, HEX_ZEBRA)
            set_cell_margins(cell, top=65, bottom=65, left=80, right=80)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            if c_idx in (0, 2):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(v)
            r.font.name = "Calibri"
            r.font.size = Pt(9.3)
            if c_idx in (0, 1):
                r.bold = True
                r.font.color.rgb = COLOR_NAVY
            else:
                r.font.color.rgb = COLOR_DARK_TEXT

    # --------------------------------------------------------------------------
    # SEÇÃO 3: ROTEIRO ALINHADO SLIDE A SLIDE (SLIDES 01 A 10)
    # --------------------------------------------------------------------------
    add_section_heading(doc, "3. ROTEIRO COMPLETO SLIDE A SLIDE", page_break_before=True)

    for idx, s in enumerate(SLIDES_ROTEIRO):
        # Quebra de página a cada 2 slides após o primeiro par para deixar cada página limpa
        if idx > 0 and idx % 2 == 0:
            doc.add_page_break()

        # Cabeçalho do Slide (barra azul-marinho compacta)
        hdr_tbl = doc.add_table(rows=1, cols=2)
        hdr_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_tbl.autofit = False
        c_left, c_right = hdr_tbl.rows[0].cells[0], hdr_tbl.rows[0].cells[1]
        c_left.width = Cm(12.2)
        c_right.width = Cm(4.8)
        set_cell_background(c_left, HEX_NAVY)
        set_cell_background(c_right, HEX_NAVY)
        set_cell_margins(c_left, top=80, bottom=80, left=140, right=80)
        set_cell_margins(c_right, top=80, bottom=80, left=80, right=140)
        c_left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        c_right.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        pl = c_left.paragraphs[0]
        pl.paragraph_format.space_after = Pt(0)
        pl.paragraph_format.keep_with_next = True
        rl1 = pl.add_run(f"SLIDE {s['num']}  |  {s['titulo']}\n")
        rl1.bold = True
        rl1.font.name = "Calibri"
        rl1.font.size = Pt(11.5)
        rl1.font.color.rgb = COLOR_WHITE
        rl2 = pl.add_run(f"{s['eixo']}   •   Voz Principal: {s['lider']}")
        rl2.font.name = "Calibri"
        rl2.font.size = Pt(8.8)
        rl2.font.color.rgb = RGBColor(226, 183, 104)

        pr = c_right.paragraphs[0]
        pr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pr.paragraph_format.space_after = Pt(0)
        pr.paragraph_format.keep_with_next = True
        rr1 = pr.add_run(f"⏱ {s['cronometro']}\n")
        rr1.bold = True
        rr1.font.name = "Calibri"
        rr1.font.size = Pt(11)
        rr1.font.color.rgb = RGBColor(247, 223, 165)
        rr2 = pr.add_run(f"Duração: {s['tempo_duracao']}")
        rr2.font.name = "Calibri"
        rr2.font.size = Pt(9)
        rr2.font.color.rgb = RGBColor(203, 213, 225)

        # Linha de ancoragem na tela + ação interativa
        p_tela = doc.add_paragraph()
        p_tela.paragraph_format.space_before = Pt(5)
        p_tela.paragraph_format.space_after = Pt(2)
        p_tela.paragraph_format.line_spacing = 1.18
        p_tela.paragraph_format.keep_with_next = True
        add_formatted_runs(p_tela, f"**Na Tela:** {s['tela_resumo']}", base_size=9.5, base_color=COLOR_SLATE)

        p_act = doc.add_paragraph()
        p_act.paragraph_format.space_before = Pt(0)
        p_act.paragraph_format.space_after = Pt(4)
        p_act.paragraph_format.keep_with_next = True
        add_formatted_runs(p_act, f"**Ação no HTML:** {s['acao_tela']}", base_size=9.5, base_color=COLOR_GOLD_DARK)

        # Tópicos de segurança (consulta rápida)
        for b in s["bullets"]:
            pb = doc.add_paragraph(style="List Bullet")
            pb.paragraph_format.space_before = Pt(0)
            pb.paragraph_format.space_after = Pt(1.5)
            pb.paragraph_format.line_spacing = 1.15
            add_formatted_runs(pb, b, base_size=9.8, base_color=COLOR_DARK_TEXT)

        # Caixa de fala sugerida (texto corrido calibrado)
        word_count = len(re.sub(r"\[.*?\]", "", s["fala"]).replace("**", "").replace("*", "").split())
        add_left_border_box(
            doc,
            title=f"LOCUÇÃO SUGERIDA — SLIDE {s['num']} ({word_count} palavras | ~{s['tempo_duracao']})",
            text_with_bold=s["fala"],
            bg_hex=HEX_GOLD_BG,
            border_hex=HEX_BORDER_GOLD,
            space_after=10,
        )

    # --------------------------------------------------------------------------
    # SEÇÃO 4: GUIA DE BOLSO PARA ARGUIÇÃO DA BANCA
    # --------------------------------------------------------------------------
    add_section_heading(doc, "4. GUIA DE BOLSO PARA POSSÍVEIS PERGUNTAS DA BANCA", page_break_before=True)

    for q, a in PERGUNTAS_BANCA:
        pq = doc.add_paragraph()
        pq.paragraph_format.space_before = Pt(6)
        pq.paragraph_format.space_after = Pt(2)
        pq.paragraph_format.keep_with_next = True
        rq = pq.add_run(q)
        rq.bold = True
        rq.font.name = "Calibri"
        rq.font.size = Pt(11)
        rq.font.color.rgb = COLOR_NAVY

        add_left_border_box(
            doc,
            title="RESPOSTA-CHAVE SUGERIDA",
            text_with_bold=a,
            bg_hex=HEX_SLATE_BG,
            border_hex=HEX_NAVY,
            space_after=6,
        )

    for path in output_paths:
        doc.save(path)
        print(f"DOCX salvo com sucesso: {path}")


# ==============================================================================
# CONSTRUTOR DO ESPELHO EM MARKDOWN
# ==============================================================================

def build_markdown(md_path: str):
    total_words = count_spoken_words()
    lines = [
        f"# {META['titulo_doc']}",
        f"**{META['instituicao']} — {META['unidade']}**  ",
        f"{META['departamento']}  ",
        f"**Pesquisa:** *{META['titulo_pesquisa']}*  ",
        f"**Discente:** {META['discente']} | **Orientador:** {META['orientador']}  ",
        f"**Total da Locução:** {total_words} palavras (~10 min a ~135 palavras/min)",
        "",
        "---",
        "",
        "## 1. O Triálogo da Pesquisa: Autores e Brevíssimo Resumo das Obras",
        "",
    ]
    for item in AUTORES_E_OBRAS:
        lines.append(f"### {item['autor']} — *{item['papel_trialogo']}*")
        if item["nota_grafia"]:
            lines.append(f"> *{item['nota_grafia']}*")
        lines.append(f"- **Perfil:** {item['perfil']}")
        lines.append(f"- **Obra ({item['obra_titulo']}):** {item['obra_resumo']}")
        lines.append("")

    lines.append("### Síntese Comparativa do Triálogo")
    lines.append("")
    lines.append("| " + " | ".join(MATRIZ_TRIALOGO[0]) + " |")
    lines.append("| :--- | :--- | :--- | :--- |")
    for row in MATRIZ_TRIALOGO[1:]:
        lines.append("| " + " | ".join(row) + " |")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## 2. Mapa de Voo e Gestão de Tempo (00:00 – 10:00)")
    lines.append("")
    lines.append("| Slide | Cronômetro | Duração | Título | Voz Principal |")
    lines.append("| :---: | :---: | :---: | :--- | :--- |")
    for s in SLIDES_ROTEIRO:
        lines.append(f"| {s['num']} | {s['cronometro']} | {s['tempo_duracao']} | {s['titulo']} | {s['lider']} |")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## 3. Roteiro Completo Slide a Slide")
    lines.append("")
    for s in SLIDES_ROTEIRO:
        lines.append(f"### Slide {s['num']} — {s['titulo']} (`{s['cronometro']}` | {s['tempo_duracao']})")
        lines.append(f"- **Eixo / Voz:** {s['eixo']} — *{s['lider']}*")
        lines.append(f"- **Na Tela:** {s['tela_resumo']}")
        lines.append(f"- **Ação no HTML:** {s['acao_tela']}")
        lines.append("- **Tópicos de Segurança:**")
        for b in s["bullets"]:
            lines.append(f"  - {b}")
        lines.append("")
        lines.append(f"> **Fala Sugerida:** {s['fala']}")
        lines.append("")

    lines.append("---")
    lines.append("")
    lines.append("## 4. Guia de Bolso para Arguição da Banca")
    lines.append("")
    for q, a in PERGUNTAS_BANCA:
        lines.append(f"**{q}**  ")
        lines.append(f"{a}")
        lines.append("")

    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"Markdown salvo com sucesso: {md_path}")


if __name__ == "__main__":
    base_dir = r"C:\Biblioteca_Pesquisa_Central\Projeto de Iniciação"
    slides_dir = os.path.join(base_dir, "SLIDES")

    docx_slides = os.path.join(slides_dir, "ROTEIRO_APRESENTACAO_10MIN_GOLEM.docx")
    docx_root = os.path.join(base_dir, "ROTEIRO_APRESENTACAO_10MIN_GOLEM.docx")
    md_slides = os.path.join(slides_dir, "ROTEIRO_APRESENTACAO_10MIN_GOLEM.md")

    build_docx([docx_slides, docx_root])
    build_markdown(md_slides)
    print(f"Contagem total de palavras faladas: {count_spoken_words()}")
