# BOA: Melhor Acessibilidade do Office

O BOA é uma poderosa suíte de melhorias de acessibilidade para o Microsoft Office, projetada para aprimorar consideravelmente a experiência com leitores de tela para usuários do NVDA. Ele corrige diretamente componentes de interface de usuário inacessíveis e introduz ferramentas de navegação rápida para o Excel e o PowerPoint.

---

## ⌨️ Referência de Teclas de Atalho

| Recurso | Combinação de Teclas | Contexto / Notas |
| :--- | :--- | :--- |
| **Entrar no Modo de Comando** | `NVDA+E` | Ativa o Modo de Prefixo de Comando (emite um bipe agudo) |
| **Analisar Layout da Planilha** | `NVDA+E`, depois `L` | Executar dentro do Excel antes de navegar pelos blocos de dados |
| **Pular para o Bloco de Dados Mais Próximo** | `NVDA+E`, depois `J` /  | Requer Análise de Layout primeiro |
| **Abrir Organizador de Planilhas em Massa** | `NVDA+E`, depois `X` | Abre a caixa de diálogo de reordenação de planilhas acessível |
| **Mover Planilha Ativa para a Esquerda** | `NVDA+Shift+SetaEsquerda` | Move a planilha ativa uma posição para cima|
| **Mover Planilha Ativa para a Direita** | `NVDA+Shift+SetaDireita` | Move a planilha ativa uma posição para baixo|
| **Mover Planilha para Início/Fim** | `NVDA+Shift+Início` / `Fim` | Envia a planilha para os limites absolutos |
| **Formatação Condicional Detalhada**| `NVDA+E`, depois `F` | Anuncia detalhes completos de formatação da célula focada |
| **Mapear Célula para Slot de Memória** | `NVDA+E`, depois `Shift+1` até `Shift+9` | Atribui a célula atual a um slot de monitoramento em segundo plano |
| **Ler Slot de Célula Monitorada** | `NVDA+E`, depois `1` até `9` | Recupera e lê o valor do slot atribuído |
| **Alternar Monitoramento em Segundo Plano** | `NVDA+E`, depois `M` | Alterna manualmente o rastreamento de cálculo em segundo plano |
| **Limpar Todos os Slots de Memória** | `NVDA+E`, depois `Backspace` | Elimina todos os monitores de células salvos em segundo plano |
| **Cancelar Modo de Comando** | `Escape` | Sai do Modo de Prefixo de Comando |

---

## 🚀 Recursos

### Melhorias no Excel

#### 1. Analisador de Layout de Planilha e Cache
Escaneie instantaneamente qualquer planilha do Excel para entender sua estrutura, elementos ocultos e blocos de dados.
* **Como funciona:** O BOA rapidamente escaneia a planilha e anuncia os blocos de dados ativos. Ele também o avisa sobre **Guias de Planilha Ocultas**, **Filtros** ativos, **Modos Protegidos** e **Limites Ocultos** (por exemplo, se colunas perto da borda direita da planilha estiverem ocultas, evitando que você perca dados fora da tela).
* **Navegação de Dados:** Após o escaneamento, você pode usar as teclas de atalho de salto do bloco de dados para alternar instantaneamente o cursor entre os blocos de dados descobertos, ignorando milhares de células vazias sem esforço.

#### 2. Organizador de Planilhas em Massa
Reordene e organize instantaneamente várias planilhas de uma vez usando uma caixa de diálogo totalmente acessível.
* **Como funciona:** Abre uma caixa de diálogo onde você pode selecionar uma planilha e mapeá-la para uma nova posição. Os movimentos agendados são listados em uma tabela de dados (pressione `Del` para remover um erro). Clique em `OK` e sua pasta de trabalho será reorganizada instantaneamente.

#### 3. Movimentador Rápido de Planilhas
Mova a planilha ativa para a esquerda, para a direita, para o início ou para o fim instantaneamente usando atalhos de teclado.

#### 4. Renomeação Acessível de Planilhas
* Ao renomear uma planilha, o NVDA nativamente tem dificuldade em ler os caracteres que você está digitando.
* O BOA injeta uma classe personalizada `ExcelSheetRenameEdit` que usa o motor `SafeRichEdit`, o que significa que você pode ler precisamente por caractere, palavra ou linha enquanto renomeia. Isso serve como um aprimoramento para o comportamento de renomeação padrão existente.

#### 5. Rastreador de Linhas/Colunas Ocultas
* Rastreia proativamente seu movimento na grade para evitar que você perca dados ocultos ou filtrados.
* **Células Fragmentadas Cruzadas:** Se você pular uma seção fortemente fragmentada ou oculta da grade (por exemplo, movendo-se da Linha 3 para a Linha 10 porque as Linhas 4 a 9 estão ocultas), o BOA anuncia explicitamente "Linhas 4 até 9 ocultas". Isso garante que você sempre saiba quando os dados foram ignorados na estrutura.

#### 6. Anunciador de Formatação Condicional
* Lê automaticamente a cor, estilo da fonte e sombreamento de fundo de células que foram alteradas dinamicamente pelas regras de Formatação Condicional do Excel.
* Dá a você o verdadeiro estado visual da célula em vez do valor subjacente. Inicialmente, ao focar na célula, ele anuncia "tem formatação condicional e alguns outros detalhes menores". Para informações abrangentes, use a configuração detalhada de tecla de atalho, que é NVDA E e F.

#### 7. Melhor anúncio de seleção
lê se a célula ou intervalo está selecionado ou desmarcado.

#### 8 Monitor de Células:
* **Monitor de Células:** Use caminhos de comando para mapear células específicas em slots de memória. Você pode voltar e lê-las a qualquer momento usando o slot numérico atribuído.
* **Monitoramento Contínuo:** Células com slots são automaticamente monitoradas em segundo plano. Se o Excel acionar um recálculo ou edição de célula, o BOA anuncia instantaneamente o novo valor. Alterne manualmente ou limpe todos através dos slots de comando.

### Melhorias no PowerPoint

#### 1. Seletores de Cor Acessíveis
* Desbloqueia a caixa de diálogo Cor Personalizada no PowerPoint.
* Identifica e lê explicitamente as caixas de edição "Vermelho", "Verde" e "Azul" corretamente (substituindo `PowerPointRGBEdit`).
* Mapeia o campo de entrada Hex anteriormente invisível para que o NVDA possa ler o valor completo da cor Hex de forma clara.

#### 2. Suporte à Grade de Cores Padrão
* Navegar na grade hexagonal de cores "Padrão" do PowerPoint normalmente é lido como "Gráfico" ou silêncio.
* O BOA rastreia suas teclas de seta através do hexágono e busca silenciosamente o valor da cor oculta, anunciando-o para você em tempo real (por exemplo, "Cor #FF0000").

### Infraestrutura e Mecanismos Técnicos

#### O Modo de Prefixo de Comando
Para evitar conflitos de teclas com outros plugins do NVDA, o BOA usa um **Modo de Prefixo de Comando**:
1. Pressione a tecla de atalho de ativação para entrar no Modo de Comando. Você ouvirá um bipe agudo.
2. Pressione uma tecla secundária para acionar um recurso específico.
3. Se você pressionar uma tecla inválida, ouvirá um bipe de erro.

#### Personalização e Painel de Configurações
* Os recursos do BOA são totalmente modulares e podem ser ativados ou desativados a qualquer momento. Vá para `Menu do NVDA -> Preferências -> Configurações -> Melhorias do Office BOA` para ativar ou desativar recursos individuais.
* **Teclas de Aceleração Inteligentes:** Cada configuração possui um atalho matematicamente exclusivo `Alt+Tecla` dentro do painel. Por exemplo, pressione `Alt+E` para ir instantaneamente para o grupo do Excel, `Alt+P` para PowerPoint e `Alt+W` para Word.
* As configurações são salvas de forma segura em um arquivo JSON independente (`boa_settings.json`), garantindo que sua configuração principal do NVDA nunca seja corrompida.
* Se o Microsoft Office corrigir oficialmente um bug de acessibilidade no futuro, você pode desativar o gancho específico do BOA com segurança sem perder o resto da funcionalidade do add-on.

#### Segurança e Limites de Integração
* Injeções da área de transferência verificam rigorosamente os IDs de processo em primeiro plano da janela para evitar o vazamento de dados para outros aplicativos.
* algumas teclas de atalho personalizadas são totalmente expostas na caixa de diálogo Definir Comandos do NVDA, sob a categoria "Better Office Accessibility".

---

## 📋 Requisitos

* **NVDA:** Versão 2026.1.0 ou superior.
* **Aplicativos:** Microsoft Excel e Microsoft PowerPoint.

---

## 💾 Instalação

1. Baixe o último arquivo de versão `.nvda-addon`, ou localize-o na loja nativa de Add-ons do NVDA.
2. se instalar a partir de um arquivo, Abra o arquivo ou use `Loja de Add-ons do NVDA -> Instalar a partir de um arquivo externo`.
3. Reinicie o NVDA.

---

## 🛠️ Log de Alterações

### v1.6.0
* **Suporte Abrangente a Traduções**: O add-on está agora totalmente localizado com suporte para 17 idiomas globais. 
  *(Nota: Estas traduções foram geradas por IA, portanto alguns pequenos erros de tradução ou imprecisões podem estar presentes.)*
* **Governança Rigorosa de Código**: Aplicados cabeçalhos de direitos autorais GPL-2.0 em toda a base de código.

### Versão 1.5.0 
#### Novos Recursos
##### Radar de Fim de Dados
Ao navegar por planilhas grandes, pode ser difícil dizer se uma célula vazia significa que você chegou ao fim de uma lista ou se há simplesmente uma lacuna nos dados. O **Radar de Fim de Dados** atua como uma verificação de perímetro inteligente para evitar que você navegue às cegas por espaços vazios.
Sempre que você navega para uma célula vazia, o BOA escaneia instantaneamente as células restantes em sua direção de viagem. Se não houver absolutamente nenhum dado restante, ele anunciará proativamente:
* *"Não há mais dados abaixo"*
* *"Não há mais dados acima"*
* *"Não há mais dados à direita"*
* *"Não há mais dados à esquerda"*
**Opções de Configuração:**
Você pode configurar este recurso via `Preferências do NVDA -> Configurações -> Melhorias do Office BOA`. Como as planilhas podem conter complexidades ocultas (como fórmulas invisíveis ou linhas recolhidas), o radar oferece três modos de operação:
1. **Desligado**: Desativa o radar inteiramente.
2. **Verificação Estrita de Memória (ContA) [Padrão]**: A abordagem mais segura e rápida. Ele verifica a memória bruta da planilha. Se detectar *qualquer coisa* abaixo de você (incluindo linhas ocultas, texto, números ou fórmulas invisíveis), ele permanece completamente silencioso para evitar falsos alarmes. Ele só anuncia "Não há mais dados" quando o restante da planilha está 100% matematicamente em branco.
3. **Apenas Dados Visíveis (Motor Matemático)**: Um motor altamente avançado projetado para planilhas complexas. Ele filtra inteligentemente linhas ocultas e fórmulas invisíveis (por exemplo, `=""`). Ele só permanecerá silencioso se houver números reais ou texto visível em seu caminho.

### Versão 1.4 - 2026-06-12
#### Novos Recursos
* **Monitor de Células:** Use caminhos de comando para mapear células específicas em slots de memória. Você pode voltar e lê-las a qualquer momento usando o slot numérico atribuído.
* **Monitoramento Contínuo:** Células com slots são automaticamente monitoradas em segundo plano. Se o Excel acionar um recálculo ou edição de célula, o BOA anuncia instantaneamente o novo valor. Alterne manualmente ou limpe todos através dos slots de comando.

#### Correções de Bugs

### Versão 1.3.0 — 2026-06-05
*Lançamento final.*

#### Novos Recursos
* **Analisador de Layout da Planilha:** Adicionada infraestrutura poderosa de escaneamento de layout. Detecta instantaneamente Proteção de Planilha, Filtros de Coluna ativos, Guias de Planilha Ocultas e bordas absolutas ocultas enquanto armazena em cache blocos de dados descobertos.
* **Navegação de Bloco de Dados Guiada:** A navegação pós-análise permite saltos de cursor imediatos entre grandes grupos de dados, ignorando células vazias perfeitamente.
* **Anunciador de Formatação Condicional:** Detecta e lê automaticamente a cor visual dinâmica, estilo de fonte e sombreamento de fundo das células alteradas pelas regras de Formatação Condicional do Excel.
* **Aceleradores de Configurações Explícitas:** A GUI de Configurações do BOA foi completamente reformulada para cumprir rigorosamente com a arquitetura do NVDA. Cada caixa de seleção de recurso agora possui um atalho `Alt+Letra` globalmente exclusivo, evitando a repetição de teclado e eliminando falhas de navegação pela primeira letra.

#### Correções de Bugs
* **Detecção de Limite de Borda Absoluta:** As verificações de borda do `UsedRange` COM nativas foram substituídas por verificações de limites matemáticos 1D absolutos (`Linha 1048576` e `Coluna 16384`) para garantir a detecção de linhas/colunas ocultas mesmo que estejam muito fora do bloco de dados ativo.
* **Recuos Seguros de Propriedade COM Preguiçosa:** Loops de propriedade COM endurecidos para evitar que as threads do NVDA travem ao avaliar milhões de estruturas ocultas contíguas.

### Versão 1.2.0 — 2026-06-03
*Lançamento final.*

#### Novos Recursos
* **Cache de Início de App:** Grande reformulação arquitetônica. Os módulos principais agora são carregados sob demanda exatamente quando você foca em aplicativos do Office, eliminando atrasos de inicialização, resolvendo completamente a falha de foco em objetos "desconhecidos" nas caixas de diálogo de renomeação e preservando a estrutura do código-fonte com vários arquivos.
* **Rastreador de Células Aprimorado (Matemática COM 1D):** Reescreveu a lógica de detecção de espaço em células ocultas para avaliar apenas seções transversais unidimensionais (`current_col` ou `current_row`). Isso reduz a carga de cálculo do COM em mais de 16 milhões de células, eliminando instantaneamente os travamentos de navegação ao pular intervalos ocultos.
* **Limpeza de Memória do Processo:** Implementou o rastreamento do Window Handle do Excel (`Hwnd`) para detectar quando o usuário fecha e reabre o Excel. Isso limpa ativamente a memória do estado global obsoleto e resolve completamente o falso anúncio de "Planilha oculta" ao abrir um novo "Book1".

#### Correções de Bugs
* **Duplo Anúncio de Seleção:** Migrou de `winUser.getKeyState` assíncrono não confiável e implementou `api.getLastInputGesture()` para suprimir perfeitamente anúncios duplos ao usar as teclas Shift+Seta.
* **Desativação do Detector de Limites:** O Detector de Limites Proativo foi desativado para proteger a estabilidade da navegação nativa do NVDA, recuando inteiramente para o rastreador de saltos de lacunas.

### Versão 1.1.0 — 2026-05-30
*Lançamento final.*

#### Novos Recursos
* **GUI de Configurações:** Adicionado um painel nativo de Melhorias do Office BOA dentro de `NVDA -> Preferências -> Configurações` para alternar facilmente os recursos ativados ou desativados.
* **Gancho SafeRichEdit:** Evita travamentos silenciosos do NVDA ao interagir com controles RichEdit no Office 2024.
* **Teclas de Atalho Personalizáveis:** Todas as teclas de atalho do BOA agora estão totalmente expostas na caixa de diálogo Definir Comandos do NVDA na categoria "Better Office Accessibility".
* **Excel: Detecção de Salto de Linha/Coluna Oculta:** Anuncia proativamente ao navegar além de linhas ou colunas ocultas, garantindo que você nunca perca dados filtrados. Pode ser alternado nas configurações.

#### Correções de Bugs
* **Segurança de Thread:** Removeu todos os atrasos de bloqueio (`time.sleep`) e os substituiu por retornos de chamada assíncronos não bloqueantes do NVDA para garantir que o leitor de tela nunca gagueje durante operações em segundo plano.

### Versão 1.0.0 — 2026-05-24
*Lançamento inicial público.*

#### Novos Recursos
* **Excel: Organizador de Planilhas em Massa:** Reordene instantaneamente várias planilhas de uma vez usando uma caixa de diálogo totalmente acessível.
* **Excel: Movimentador Rápido de Planilha:** Mova a planilha ativa para a esquerda, direita, início ou fim por meio de comandos de teclado.
* **Excel: Renomeação Acessível de Planilha:** Intercepta o campo de renomeação nativo inacessível e o substitui por uma caixa de diálogo acessível confiável.
* **Excel: Rastreamento Inteligente de Seleção:** Anuncia com precisão seleções e desmarcações de intervalos de múltiplas células.
* **PowerPoint: Seletores de Cor Acessíveis:** Permite que o NVDA leia com precisão os valores RGB e Hex dentro da caixa de diálogo Cor Personalizada.
* **PowerPoint: Suporte à Grade de Cores Padrão:** Intercepta a navegação com teclas de seta para ler os códigos Hex ocultos da grade de cores hexagonais inacessíveis.
