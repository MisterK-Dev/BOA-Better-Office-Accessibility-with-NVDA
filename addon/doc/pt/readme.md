# BOA: Better Office Accessibility

O BOA é um poderoso conjunto de melhorias de acessibilidade para o Microsoft Office, projetado para aprimorar significativamente a experiência do leitor de tela para usuários do NVDA. Ele corrige diretamente componentes de interface inacessíveis e introduz ferramentas de navegação rápida para o Excel e o PowerPoint.

---

## ⌨️ Referência de Teclas de Atalho

| Recurso | Combinação de Teclas | Contexto / Notas |
| :--- | :--- | :--- |
| **Entrar no Modo de Comando** | `[Prefix]` (Padrão: `NVDA+E`) | Ativa o Modo de Prefixo de Comando (emite um bipe agudo) |
| **Cancelar o Modo de Comando** | `Escape` | Sai do Modo de Prefixo de Comando |
| **MELHORIAS NO EXCEL** | | |
| **Analisar o Layout da Planilha** | `[Prefix]`, depois `L` | Executar no Excel antes de navegar pelos blocos de dados |
| **Saltar para o Bloco de Dados Mais Próximo** | `[Prefix]`, depois `J` | Requer análise de layout prévia |
| **Abrir Organizador em Massa de Planilhas** | `[Prefix]`, depois `X` | Abre o diálogo acessível de reordenação de planilhas |
| **Anunciador de Fórmula Pura** | `[Prefix]`, depois `F2` | Toque único para ouvir a string da fórmula pura |
| **Editor de Fórmulas Avançado** | `[Prefix]`, depois `F2` duas vezes | Toque duplo para abrir o editor de fórmulas multilinha acessível |
| **Rastrear Precedentes** | `[Prefix]`, depois `Shift+P` | Rastrear precedentes, o mesmo recurso de maneira acessível. |
| **Rastrear Dependentes** | `[Prefix]`, depois `Shift+D` | Rastrear dependentes, o mesmo recurso de maneira acessível; pressionar Enter em uma célula irá teletransportar você para ela. |
| **Formatação Condicional Detalhada** | `[Prefix]`, depois `F` | Anuncia detalhes de formatação completos da célula em foco |
| **Mover Planilha Ativa para a Esquerda** | `NVDA+Shift+LeftArrow` | Desloca a planilha ativa uma posição para a esquerda |
| **Mover Planilha Ativa para a Direita** | `NVDA+Shift+RightArrow` | Desloca a planilha ativa uma posição para a direita |
| **Mover Planilha para o Início/Fim** | `NVDA+Shift+Home` / `End` | Envia a planilha para os limites absolutos |
| **Ocultar / Reexibir Linha** | `Ctrl+9` / `Ctrl+Shift+9` | Atalho nativo; o BOA anuncia explicitamente a alteração de visibilidade |
| **Ocultar / Reexibir Coluna** | `Ctrl+0` / `Ctrl+Shift+0` | Atalho nativo; o BOA anuncia explicitamente a alteração de visibilidade |
| **Reexibir Coluna (Alternativo)** | `NVDA+Ctrl+Shift+0` | Ignora conflitos de atalhos de idioma de entrada do Windows |
| **Mapear Célula para Slot de Memória** | `[Prefix]`, depois `Shift+1` a `Shift+9` | Atribui a célula atual a um slot de monitoramento em segundo plano |
| **Ler Slot de Célula Monitorada** | `[Prefix]`, depois `1` a `9` | Recupera e lê o valor do slot atribuído |
| **Pulo Direto para Slot** | `Alt` + `1` a `9` | Salta instantaneamente o cursor para um slot monitorado |
| **Retornar para a Célula Anterior** | `[Prefix]`, depois `\` | Teletransporta você instantaneamente de volta após verificar um slot |
| **Diálogo do Gerenciador de Slots** | `[Prefix]`, depois `Alt+M` | Abre um diálogo para visualizar e gerenciar todos os monitores ativos |
| **Alternar Monitoramento em Segundo Plano** | `[Prefix]`, depois `M` | Alterna manualmente o rastreamento de cálculos em segundo plano |
| **Limpar Todos os Slots de Memória** | `[Prefix]`, depois `Backspace` | Limpa todos os monitores de células salvos em segundo plano |
| **MELHORIAS NO POWERPOINT** | | |
| **Analisador de Layout do Slide** | `[Prefix]`, depois `L` | Analisa e anuncia o layout espacial do slide atual |
| **Analisador de Documento** | `[Prefix]`, depois `D` | Gera um Índice analítico abrangente e um relatório de integridade |
| **Organizador em Massa de Slides** | `[Prefix]`, depois `X` | Abre o diálogo acessível para reordenar múltiplos slides |
| **MELHORIAS NO WORD** | | |
| **Auditor de Formatação** | `[Prefix]`, depois `F` | Audita o documento atual para inconsistências de formatação |
| **Analisador de Documento** | `[Prefix]`, depois `D` | Analisa o layout e a estrutura do documento do Word atual |

---

## 🚀 Recursos

### Melhorias no Excel

#### 1. Analisador de Layout da Planilha e Cache
Verifique instantaneamente qualquer planilha do Excel para entender sua estrutura, elementos ocultos e blocos de dados.
* **Como funciona:** O BOA varre rapidamente a planilha e anuncia os blocos de dados ativos. Ele também avisa sobre **Guias de Planilha Ocultas**, **Filtros** ativos, **Modos Protegidos** e **Limites Externos Ocultos** (por exemplo, se as colunas próximas à borda direita da planilha estiverem ocultas, evitando que você perca dados fora da tela).
* **Navegação de Dados:** Após a varredura, você pode usar as teclas de atalho de salto de bloco de dados para mover instantaneamente o cursor entre os blocos de dados descobertos, ignorando sem esforço milhares de células vazias.

#### 2. Organizador em Massa de Planilhas
Reordene e organize instantaneamente várias planilhas de uma vez usando um diálogo totalmente acessível.
* **Como funciona:** Abre um diálogo onde você pode selecionar uma planilha e mapeá-la para uma nova posição. Os movimentos agendados são listados em uma tabela de dados (pressione `Del` para remover um erro). Clique em `OK` e sua pasta de trabalho será rearranjada instantaneamente.

#### 3. Movimentador Rápido de Planilha
Mova a planilha ativa para a esquerda, para a direita, para o início ou para o final instantaneamente usando seus atalhos de teclado.

#### 4. Renomeação Acessível de Planilha
* Ao renomear uma planilha, o NVDA nativamente tem dificuldade para ler os caracteres que você está digitando.
* O BOA injeta uma classe personalizada `ExcelSheetRenameEdit` que usa o motor `SafeRichEdit`, o que significa que você pode ler com precisão por caractere, palavra ou linha ao renomear. Isso serve como uma melhoria ao comportamento padrão de renomeação existente.

#### 5. Rastreador de Linhas/Colunas Ocultas
* Rastreia proativamente seu movimento pela grade para evitar que você perca dados ocultos ou filtrados.
* **Cruzamento de Células Fragmentadas:** Se você saltar por uma seção altamente fragmentada ou oculta da grade (por exemplo, movendo da Linha 3 para a Linha 10 porque as Linhas 4–9 estão ocultas), o BOA anunciará explicitamente "Linhas 4 a 9 ocultas". Isso garante que você sempre saiba quando dados foram pulados na estrutura.

#### 6. Anunciador de Formatação Condicional
* Lê automaticamente a cor, o estilo da fonte e o sombreamento de fundo de células que foram alteradas dinamicamente pelas regras de Formatação Condicional do Excel.
* Fornece o real estado visual da célula em vez de apenas o valor puro subjacente. Inicialmente, ao focar na célula, ele anuncia "tem formatação condicional, e alguns outros detalhes menores". Para informações abrangentes, use a configuração detalhada de teclas de atalho, que é `NVDA+E` e depois `F`.

#### 7. Melhor anúncio de seleção
lê se a célula ou o intervalo foi selecionado ou desmarcado.

#### 8 Monitor de célula:
* **Monitor de Célula:** Use caminhos de comando para mapear células específicas para slots de memória. Você pode retornar e lê-las a qualquer momento usando o slot numérico atribuído.
* **Monitoramento Contínuo:** As células associadas aos slots são monitoradas automaticamente em segundo plano. Se o Excel acionar um recálculo ou edição de célula, o BOA anuncia instantaneamente o novo valor. Alterne manualmente ou limpe tudo através dos slots de comando.
* **Excel: Atualizações do Monitor de Célula Pro:**
  - **Diálogo do Gerenciador de Slots (`NVDA+E`, depois `Alt+M`):** Abre um diálogo listando todas as células ativamente monitoradas. Pressione `Enter` para saltar instantaneamente para uma delas.
  - **Retornar (`NVDA+E`, depois `\`):** Teletransporta você instantaneamente de volta para a célula de trabalho anterior após verificar um slot.
  - **Salto Direto para Slot (`Prefix + Alt` + `Número do Slot`):** Ignora o prefixo e salta instantaneamente para um slot de célula atribuído.

#### 9 Editor avançado
* **Excel: O Editor Avançado (Editor de Fórmulas Acessível):** Um divisor de águas absoluto para modificar fórmulas gigantescas.
  - **Toque Único `NVDA+E`, depois `F2`:** Anuncia instantaneamente a string da fórmula pura da célula ativa (ou anuncia "Nenhuma fórmula").
  - **Toque Duplo `NVDA+E`, depois `F2`:** Abre um editor multilinha totalmente acessível para modificar com segurança fórmulas gigantescas e aninhadas. O `Enter` nativo adiciona quebras de linha para facilitar a leitura, e `Ctrl+Enter` salva de volta no Excel.
  - *Verificações de Segurança:* Captura com segurança erros de sintaxe antes que eles corrompam sua planilha, e detecta erros pós-cálculo (como `#NAME?` ou `#DIV/0!`) para avisá-lo instantaneamente se uma fórmula quebrar.

#### 10 Melhorias na auditoria e avaliação de fórmulas:
* **Excel: Auditoria e Avaliação de Fórmulas:** Adicionados atalhos personalizados (`NVDA+E`, depois `Shift+P` e `NVDA+E`, depois `Shift+D`) para rastrear com segurança Precedentes e Dependentes. Além disso, o diálogo nativo "Avaliar Fórmula" do Excel agora é totalmente acessível; o NVDA lê automaticamente os resultados avaliados conforme você avança no cálculo!

### Melhorias no PowerPoint

#### 1. Seletores de Cores Acessíveis
* Desbloqueia o diálogo de Cores Personalizadas no PowerPoint.
* Identifica e lê explicitamente as caixas de edição de "Vermelho", "Verde" e "Azul" corretamente (sobrepondo `PowerPointRGBEdit`).
* Mapeia o campo de entrada Hex anteriormente invisível para que o NVDA possa ler o valor de cor Hex completo de forma limpa.

#### 2. Suporte à Grade de Cores Padrão
* A navegação na grade hexagonal de cores "Padrão" do PowerPoint normalmente é lida como "Gráfico" ou silêncio.
* O BOA rastreia suas teclas de seta pelo hexágono e busca silenciosamente o valor de cor oculto, anunciando-o para você em tempo real (por exemplo, "Cor #FF0000").

#### 3 Organizador em Massa de Slides:
* **PowerPoint: Organizador em Massa de Slides (Experimental) (`NVDA+E`, depois `X`):** Semelhante ao recurso do Excel, agora você pode reordenar, mover e organizar instantaneamente múltiplos slides do PowerPoint de uma vez usando um diálogo totalmente acessível.

#### 4 Analisador de layout de slide
* **PowerPoint: Analisador de Layout do Slide (Experimental) (`NVDA+E`, depois `L`):** Varre instantaneamente o slide ativo para entender seu layout espacial e restrições de acessibilidade, garantindo uma experiência de leitor de tela totalmente fluida e responsiva. Ou seja, aqui você obterá detalhes sobre o slide atual semelhante ao analisador de layout de planilha do Excel.

#### 5 Analisador Completo de Documento [PPT]
* **PowerPoint: Analisador Completo de Documento (Experimental) (`NVDA+E`, depois `D`):** Uma ferramenta de acessibilidade altamente avançada, processada em segundo plano, que mapeia uma apresentação inteira sem travar o sintetizador de voz do NVDA. Ela fornece um Índice Virtual profundamente navegável, detecta Incompatibilidades de Ordem de Leitura (Ordem Visual vs. Ordem Z), sinaliza slides com "Muralha de Texto" (excesso de texto) e mapeia objetos complexos como SmartArt e Tabelas de Dados.

#### 6 Melhorias no movimento [ajuste] de formas:
* **PowerPoint: Modo de Áudio de Movimento de Forma (Experimental):** Introduz dicas de Áudio Espacial 3D à tela do PowerPoint. Fornece feedback auditivo indicando a direção e os limites de um objeto conforme você o move, melhorando significativamente a percepção espacial.

### Melhorias no Word:
#### 1. Analisador de Documento inspirado e derivado do complemento word access de Paul:
* **Word: Analisador de Documento (`NVDA+E`, depois `D`):** Exibe instantaneamente uma visão geral estrutural do seu documento do Word. *(Uma nota especial de crédito e agradecimento ao Paul: Este recurso foi diretamente inspirado por seu brilhante complemento "Word Access". Somos profundamente gratos por seu trabalho fundamental nesta área!)*

#### 2 Auditor de Formatação
* **Word: Auditor de Formatação (`NVDA+E`, depois `F`):** Varre seu documento do Word em busca de inconsistências de formatação para garantir padrões visuais.

#### 3 Leitor de notas de rodapé:
* **Word: Anunciador Automático de Notas de Rodapé:** As notas de rodapé serão agora anunciadas automaticamente integradas no texto conforme você lê, dependendo das suas configurações personalizadas do BOA. *(Nota: O suporte para notas de fim e comentários está planejado para uma versão futura).*

### Infraestrutura e Mecanismos Técnicos

#### O Modo de Prefixo de Comando
Para evitar conflitos de teclas de atalho com outros complementos do NVDA, o BOA usa um **Modo de Prefixo de Comando**:
1. Pressione a tecla de atalho de ativação para entrar no Modo de Comando. Você ouvirá um bipe agudo. O padrão é NVDA mais E.
2. Pressione uma tecla secundária para acionar um recurso específico.
3. Se você pressionar uma tecla inválida, ouvirá um bipe de erro.

#### Painel de Configurações e Personalização
* Os recursos do BOA são totalmente modulares e podem ser ativados ou desativados a qualquer momento. Vá em `Menu do NVDA -> Preferências -> Configurações -> BOA Office Enhancements` para ligar ou desligar recursos individuais.
* **Teclas de Atalho Inteligentes:** Cada configuração apresenta um atalho acelerador `Alt+Tecla` matematicamente exclusivo dentro do painel. Por exemplo, pressione `Alt+E` para saltar instantaneamente para o grupo do Excel, `Alt+P` para o PowerPoint e `Alt+W` para o Word.
* As configurações são salvas com segurança em um arquivo JSON independente (`boa_settings.json`), garantindo que sua configuração principal do NVDA nunca seja corrompida.
* Se o Microsoft Office corrigir oficialmente um bug de acessibilidade no futuro, você poderá desativar com segurança o gancho de substituição específico do BOA sem perder o restante da funcionalidade do complemento.
* **Personalização de Gestos de Entrada:** Todos os recursos em todos os aplicativos do Office foram expostos explicitamente ao diálogo nativo de Gestos de Entrada do NVDA, concedendo a você total liberdade para personalizar cada atalho de teclado.

#### Segurança e Limites de Integração
* As injeções na área de transferência verificam rigorosamente os IDs de processos de primeiro plano da janela para evitar o vazamento de dados para outros aplicativos.
* algumas teclas de atalho personalizadas estão totalmente expostas no diálogo Gestos de Entrada do NVDA sob a categoria "Better Office Accessibility".

---

## 📋 Requisitos

* **NVDA:** Versão 2026.1.0 ou posterior.
* **Aplicativos:** Microsoft Excel e Microsoft PowerPoint.

---

## 💾 Instalação

1. Baixe o arquivo de lançamento `.nvda-addon` mais recente ou localize-o na Loja de Complementos nativa do NVDA.
2. se estiver instalando a partir de um arquivo, abra o arquivo ou use a `Loja de Complementos do NVDA -> Instalar de arquivo externo`.
3. Reinicie o NVDA.

---

## 🛠️ Histórico de Alterações

### Version 2.0.1
#### Melhorias de UX/UI
* **Diálogo de configurações em abas:** Reorganização do painel de configurações do BOA em abas acessíveis (&Excel, &Word e &PowerPoint) usando `wx.Notebook`, melhorando consideravelmente a navegação do leitor de tela e eliminando longas listas de rolagem. Você pode alternar rapidamente entre as abas usando `Alt+E`, `Alt+W`, `Alt+P` ou os atalhos padrão `Ctrl+PageDown`/`Ctrl+PageUp`.
* **Compatibilidade com o NVDA 2026.2:** Testado e certificado para o NVDA 2026.2.

### Versão 2.0.0
#### Novos Recursos
* **PowerPoint: Analisador Completo de Documento (Experimental) (`NVDA+E`, depois `D`):** Uma ferramenta de acessibilidade altamente avançada, processada em segundo plano, que mapeia uma apresentação inteira sem travar o sintetizador de voz do NVDA. Ela fornece um Índice Virtual profundamente navegável, detecta Incompatibilidades de Ordem de Leitura (Ordem Visual vs. Ordem Z), sinaliza slides com "Muralha de Texto" (excesso de texto) e mapeia objetos complexos como SmartArt e Tabelas de Dados.
* **PowerPoint: Analisador de Layout do Slide (Experimental) (`NVDA+E`, depois `L`):** Varre instantaneamente o slide ativo para entender seu layout espacial e restrições de acessibilidade, garantindo uma experiência de leitor de tela totalmente fluida e responsiva. Ou seja, aqui você obterá detalhes sobre o slide atual semelhante ao analisador de layout de planilha do Excel.
* **PowerPoint: Organizador em Massa de Slides (Experimental) (`NVDA+E`, depois `X`):** Semelhante ao recurso do Excel, agora você pode reordenar, mover e organizar instantaneamente múltiplos slides do PowerPoint de uma vez usando um diálogo totalmente acessível.
* **PowerPoint: Modo de Áudio de Movimento de Forma (Experimental):** Introduz dicas de Áudio Espacial 3D à tela do PowerPoint. Fornece feedback auditivo indicando a direção e os limites de um objeto conforme você o move, melhorando significativamente a percepção espacial. Como mencionado, isto é experimental, aguardando feedbacks para melhorias.
* **Word: Auditor de Formatação (`NVDA+E`, depois `F`):** Varre seu documento do Word em busca de inconsistências de formatação para garantir padrões visuais.
* **Word: Analisador de Documento (`NVDA+E`, depois `D`):** Exibe instantaneamente uma visão geral estrutural do seu documento do Word. *(Uma nota especial de crédito e agradecimento ao Paul: Este recurso foi diretamente inspirado por seu brilhante complemento "Word Access". Somos profundamente gratos por seu trabalho fundamental nesta área!)*
* **Word: Anunciador Automático de Notas de Rodapé:** As notas de rodapé serão agora anunciadas automaticamente integradas no texto conforme você lê, dependendo das suas configurações personalizadas do BOA. *(Nota: O suporte para notas de fim e comentários está planejado para uma versão futura).*
* **Excel: O Editor Avançado (Editor de Fórmulas Acessível):** Um divisor de águas absoluto para modificar fórmulas gigantescas.
  - **Toque Único `NVDA+E`, depois `F2`:** Anuncia instantaneamente a string da fórmula pura da célula ativa (ou anuncia "Nenhuma fórmula").
  - **Toque Duplo `NVDA+E`, depois `F2`:** Abre um editor multilinha totalmente acessível para modificar com segurança fórmulas gigantescas e aninhadas. O `Enter` nativo adiciona quebras de linha para facilitar a leitura, e `Ctrl+Enter` salva de volta no Excel.
  - *Verificações de Segurança:* Captura com segurança erros de sintaxe antes que eles corrompam sua planilha, e detecta erros pós-cálculo (como `#NAME?` ou `#DIV/0!`) para avisá-lo instantaneamente se uma fórmula quebrar.
* **Excel: Auditoria e Avaliação de Fórmulas:** Adicionados atalhos personalizados (`NVDA+E`, depois `Shift+P` e `NVDA+E`, depois `Shift+D`) para rastrear com segurança Precedentes e Dependentes. Além disso, o diálogo nativo "Avaliar Fórmula" do Excel agora é totalmente acessível; o NVDA lê automaticamente os resultados avaliados conforme você avança no cálculo!
* **Excel: Atualizações do Monitor de Célula Pro:**
  - **Diálogo do Gerenciador de Slots (`NVDA+E`, depois `Alt+M`):** Abre um diálogo listando todas as células ativamente monitoradas. Pressione `Enter` para saltar instantaneamente para uma delas.
  - **Retornar (`NVDA+E`, depois `\`):** Teletransporta você instantaneamente de volta para a célula de trabalho anterior após verificar um slot.
  - **Salto Direto para Slot (`Alt` + `Número do Slot`):** Ignora o prefixo completamente e salta instantaneamente para um slot de célula atribuído.
* **Personalização de Gestos de Entrada:** Todos os recursos em todos os aplicativos do Office foram expostos explicitamente ao diálogo nativo de Gestos de Entrada do NVDA, concedendo a você total liberdade para personalizar cada atalho de teclado.

#### Melhorias de UX/UI
* **Relatórios Navegáveis Unificados:** Adotamos um sistema unificado de relatórios em HTML em todo o complemento. Recursos como o Anunciador de Formatação Condicional do Excel, Analisadores de Layout e Analisadores de Documento não apenas falam blocos gigantescos de texto; seus resultados agora abrem em uma janela HTML nativa e navegável, permitindo que você revise as informações no seu próprio ritmo.
* **Excel: Rastreamento Aprimorado de Dependentes/Precedentes:** Melhorada significativamente a saída de fala para os atalhos nativos de rastreamento de fórmulas do Excel (`Ctrl+[` para Precedentes Diretos, e `Ctrl+]` para Dependentes Diretos). O NVDA agora anunciará explicitamente exatamente quais células foram selecionadas.
* **Excel: Suporte a Células Mescladas:** As células mescladas são agora detectadas corretamente e anunciadas explicitamente pelo rastreador de células que pula lacunas.

#### Correções de Bugs
* **Word: Leitura Dupla de Itens de Lista:** Implementada uma correção temporária para corrigir o bug onde o NVDA lia duas vezes os itens de lista de parágrafos em certas visualizações do Word.
* **Excel: Bug de Localização do Monitor de Célula:** Resolvidos bugs de rastreamento subjacentes causados pelas recentes updates de localização de tradução.

### O que há de novo na v1.6.1
* **Localização Profunda de Arquivos**: Corrigidas traduções de strings ausentes no interior dos módulos de melhorias do Excel (como o Analisador de Layout da Planilha e o Movimentador Rápido de Planilha) para garantir 100% de cobertura de localização.
* **Suporte de Tradução Ampliado**: Adicionados 7 novos idiomas ao sistema (turco, polonês, coreano, ucraniano, tcheco, urdu e panjabi).
  *(Nota: Estas traduções foram geradas por IA, portanto, alguns pequenos erros de tradução ou imprecisões podem estar presentes.)*

### v1.6.0
* **Suporte Amplo de Tradução**: O complemento agora está totalmente localizado com suporte para 17 idiomas globais.
  *(Nota: Estas traduções foram geradas por IA, portanto, alguns pequenos erros de tradução ou imprecisões podem estar presentes.)*
* **Governança Estrita de Código**: Aplicados cabeçalhos de direitos autorais GPL-2.0 em toda a base de código.""",

### Versão 1.5.0
#### Novos Recursos
##### End of Data Radar
Ao navegar por planilhas grandes, pode ser difícil saber se uma célula vazia significa que você chegou ao fim de uma lista ou se há simplesmente uma lacuna nos dados. O **Radar de Fim de Dados** funciona como uma verificação de perímetro inteligente para evitar que você use as setas cegamente pelo espaço vazio.
Sempre que você navega para uma célula vazia, o BOA varre instantaneamente as células restantes na sua direção de deslocamento. Se não houver absolutamente nenhum dado restante, ele anunciará proativamente:
* *"Não há mais dados abaixo"*
* *"Não há mais dados acima"*
* *"Não há mais dados à direita"*
* *"Não há mais dados à esquerda"*
**Opções de Configuração:**
Você pode configurar este recurso através de `Preferências do NVDA -> Configurações -> BOA Office Enhancements`. Como as planilhas podem conter complexidades ocultas (como fórmulas invisíveis ou linhas recolhidas), o radar oferece três modos de operação:
1. **Desativado**: Desativa o radar completamente.
2. **Verificação Estrita de Memória (CountA) [Padrão]**: A abordagem mais rápida e segura. Ela verifica a memória pura da planilha. Se detectar *qualquer coisa* abaixo de você (incluindo linhas ocultas, texto, números ou fórmulas invisíveis), permanece completamente silenciosa para evitar alarmes falsos. Ela só anuncia "Não há mais dados" quando o restante da planilha for 100% matematicamente em branco.
3. **Apenas Dados Visíveis (Motor Matemático)**: Um motor altamente avançado projetado para planilhas complexas. Ele filtra inteligentemente linhas ocultas e fórmulas invisíveis (por exemplo, `=""`). Ele só permanecerá silencioso se restarem números ou textos visíveis e reais em seu caminho.

### Versão 1.4 - 2026-06-12
#### Novos Recursos
* **Monitor de Célula:** Use caminhos de comando para mapear células específicas para slots de memória. Você pode retornar e lê-las a qualquer momento usando o slot numérico atribuído.
* **Monitoramento Contínuo:** As células associadas aos slots são monitoradas automaticamente em segundo plano. Se o Excel acionar um recálculo ou edição de célula, o BOA anuncia instantaneamente o novo valor. Alterne manualmente ou limpe tudo através dos slots de comando.

#### Correções de Bugs

### Versão 1.3.0 — 2026-06-05
*Versão final.*

#### Novos Recursos
* **Analisador de Layout da Planilha:** Adicionada uma poderosa infraestrutura de varredura de layout. Detecta instantaneamente Proteção de Planilha, Filtros de Coluna ativos, Guias de Planilha Ocultas e bordas absolutas ocultas, enquanto armazena em cache os blocos de dados descobertos.
* **Navegação Guiada por Blocos de Dados:** A navegação pós-análise permite saltos imediatos do cursor entre os principais agrupamentos de dados, ignorando células vazias perfeitamente.
* **Anunciador de Formatação Condicional:** Detecta e lê automaticamente a cor dinâmica, o estilo da fonte e o sombreamento de fundo de células alteradas pelas regras de Formatação Condicional do Excel.
* **Aceleradores de Configurações Explícitos:** Totalmente reformulada a GUI de Configurações do BOA para cumprir rigorosamente com a arquitetura do NVDA. Cada caixa de seleção de recurso agora possui um atalho `Alt+Letra` globalmente exclusivo, evitando a alternância de teclado e eliminando falhas de navegação pela primeira letra.

#### Correções de Bugs
* **Detecção de Limites de Borda Absolutos:** Substituídas as verificações de borda nativas COM `UsedRange` por verificações de limites matemáticos 1D absolutos (`Linha 1048576` e `Coluna 16384`) para garantir a detecção de linhas/colunas ocultas mesmo se estiverem muito fora do bloco de dados ativo.
* **Saídas Seguras de Propriedades COM Lazy:** Fortalecidos os loops de propriedades COM para evitar travamentos de threads do NVDA ao avaliar milhões de estruturas ocultas contíguas.

### Versão 1.2.0 — 2026-06-03
*Versão final.*

#### Novos Recursos
* **Cache de Inicialização do Aplicativo:** Grande reformulação arquitetônica. Os módulos principais agora são carregados tardiamente (lazy-loaded) exatamente quando você foca em aplicativos do Office, eliminando o atraso de inicialização, resolvendo completamente a falha de foco do objeto 'desconhecido' em diálogos de renomeação e preservando a estrutura da base de código multifivos.
* **Rastreador de Células Aprimorado (Matemática COM 1D):** Reesrita a lógica de detecção de lacunas de células ocultas para avaliar apenas seções transversais unidimensionais (`current_col` ou `current_row`). Isso reduz a carga de cálculo COM em mais de 16 milhões de células, eliminando instantaneamente travamentos de navegação ao saltar intervalos ocultos.
* **Limpeza de Memória de Processo:** Implementado rastreamento de Alça de Janela (`Hwnd`) do Excel para detectar quando o usuário fecha e reabre o Excel. Isso limpa ativamente a memória de estado global desatualizada e resolve completamente o anúncio falso de "Planilha oculta" ao abrir um novo "Pasta1".

#### Correções de Bugs
* **Anúncio Duplo de Seleção:** Migrado do não confiável assíncrono `winUser.getKeyState` e implementado `api.getLastInputGesture()` para suprimir perfeitamente anúncios duplos ao usar as teclas Shift+Seta.
* **Desativação do Detector de Limites:** O Detector de Limites Proativo foi desativado para proteger a estabilidade de navegação nativa do NVDA, recorrendo inteiramente ao rastreador de pulo de lacunas.

### Versão 1.1.0 — 2026-05-30
*Versão final.*

#### Novos Recursos
* **GUI de Configurações:** Adicionado um painel nativo BOA Office Enhancements dentro de `NVDA -> Preferências -> Configurações` para alternar facilmente os recursos entre ativado ou desativado.
* **Gancho SafeRichEdit:** Evita travamentos silenciosos do NVDA ao interagir com controles RichEdit no Office 2024.
* **Atalhos Personalizáveis:** Todos os atalhos do BOA agora estão totalmente expostos no diálogo Gestos de Entrada do NVDA sob a categoria "Better Office Accessibility".
* **Excel: Detecção de Pulo de Linha/Coluna Oculta:** Anuncia proativamente ao navegar por linhas ou colunas ocultas, garantindo que você nunca perca dados filtrados. Pode ser alternado nas configurações.

#### Correções de Bugs
* **Segurança de Threads:** Removidos todos os atrasos bloqueantes (`time.sleep`) e substituídos por retornos de chamada (callbacks) assíncronos não bloqueantes do NVDA para garantir que o leitor de tela nunca gagueje durante operações em segundo plano.

### Versão 1.0.0 — 2026-05-24
*Lançamento público inicial.*

#### Novos Recursos
* **Excel: Organizador em Massa de Planilhas:** Reordene instantaneamente várias planilhas de uma vez usando um diálogo totalmente acessível.
* **Excel: Movimentador Rápido de Planilha:** Mova a planilha ativa para a esquerda, direita, início ou fim através de comandos de teclado.
* **Excel: Renomeação Acessível de Planilha:** Intercepta o campo nativo inacessível de renomeação e o substitui por um diálogo acessível confiável.
* **Excel: Rastreamento Inteligente de Seleção:** Anuncia com precisão seleções e desmarcações de intervalos de várias células.
* **PowerPoint: Seletores de Cores Acessíveis:** Permite ao NVDA ler com precisão valores RGB e Hex dentro do diálogo de Cores Personalizadas.
* **PowerPoint: Suporte à Grade de Cores Padrão:** Intercepta a navegação por teclas de seta para ler códigos Hex ocultos da grade hexagonal de cores inacessível.
