# Deluge Shadow Analyzer - Design de Arquitetura e Refatoração

## 1. Visão Geral
O **Deluge Shadow Analyzer** é uma ferramenta de alta performance destinada à verificação de cobertura de jatos de água (sistemas de dilúvio) e análise de obstruções em modelos industriais complexos. Este documento descreve a arquitetura proposta para garantir precisão, escalabilidade e manutenibilidade.

## 2. Stack Tecnológico

A escolha das tecnologias prioriza a performance computacional e a capacidade de lidar com malhas densas (milhões de polígonos).

*   **Linguagem:** Python 3.10+
*   **Interface Gráfica (UI):** **PySide6 (Qt)**. Escolhido por ser o padrão industrial para ferramentas de engenharia, oferecendo controles nativos robustos e integração perfeita com janelas OpenGL.
*   **Motor 3D & Processamento:** **PyVista (wrapper do VTK)**. O VTK (Visualization Toolkit) é o "estado da arte" em visualização científica. O PyVista simplifica a API sem perder performance.
    *   *Motivo:* Suporte nativo a grandes datasets, estruturas de aceleração espacial (Locators) e renderização eficiente.
*   **Matemática e Malhas Auxiliares:** **NumPy** (álgebra linear) e **Trimesh** (operações rápidas de geometria e carregamento de formatos diversos).
*   **Paralelismo:** `multiprocessing` ou `concurrent.futures` para cálculos de ray-casting que não bloqueiem a UI.

## 3. Arquitetura do Código

O projeto seguirá o padrão MVC (Model-View-Controller) modificado para aplicações científicas.

### Estrutura de Diretórios
```
deluge_analyzer/
├── core/                   # Lógica de Negócio e Algoritmos
│   ├── __init__.py
│   ├── engine.py           # Gerenciador da simulação
│   ├── geometry.py         # Manipulação de malhas e cones
│   └── physics.py          # Ray-casting e detecção de colisão
├── io/                     # Entrada e Saída
│   ├── __init__.py
│   ├── loaders.py          # Leitores (FBX, OBJ, VTP, STP via gmsh)
│   └── exporters.py        # Exportadores (Navisworks Metadata)
├── ui/                     # Interface Gráfica
│   ├── __init__.py
│   ├── app.py              # MainWindow e loop principal
│   ├── viewport.py         # Widget 3D (PyVistaQt)
│   └── panels.py           # Paineis laterais de controle
└── utils/
    └── spatial.py          # Helpers para Octrees/KD-Trees
```

### Principais Classes

*   `DelugeScene`: Contém a malha do ambiente (Obstáculos) e a lista de `Nozzles` (Bicos injetores).
*   `Nozzle`: Define a origem, vetor direção, ângulo de abertura e alcance do cone.
*   `SimulationEngine`: Coordena o processo de cálculo de sombras.

## 4. Algoritmos e Otimização

### 4.1 Processamento de Geometria e Estruturas Espaciais
Para evitar verificar cada raio contra cada triângulo (complexidade $O(R \times T)$), utilizaremos **Estruturas de Aceleração Espacial**.
*   **Algoritmo:** Utilização do `vtkCellLocator` ou `vtkOBBTree` (Oriented Bounding Box Tree).
*   **Funcionamento:** O modelo 3D é subdividido em uma árvore espacial. O teste de interseção descarta rapidamente grandes regiões do espaço que não são tocadas pelo raio.
*   **Ganho:** Redução da complexidade para aproximadamente $O(R \times \log T)$.

### 4.2 Cálculo de Sombras (Ray Casting Otimizado)
A simulação de um cone de dilúvio será discretizada via Ray Casting estocástico ou uniforme.

1.  **Geração de Raios:** Para cada `Nozzle`, geramos $N$ vetores (ex: 10.000) distribuídos uniformemente dentro do ângulo sólido do cone.
2.  **Batch Intersection:**
    *   Utilizamos a função `locator.IntersectWithLine(p1, p2, points, cell_ids)` do VTK.
    *   Se um raio atinge uma face antes de atingir seu alcance máximo, o ponto de impacto é marcado como "Molhado".
    *   Regiões "atrás" do impacto (na mesma direção) são implicitamente sombreadas (o raio não chega lá).
3.  **Mapeamento na Malha (Coloring):**
    *   Em vez de apenas pintar os pontos de impacto, projetamos o estado (Molhado/Seco) nos vértices da malha original.
    *   Utilizamos `point_scalars` do VTK.
    *   *Azul (0.0 -> 1.0):* Vértices próximos a impactos.
    *   *Vermelho (0.0):* Vértices não atingidos.

### 4.3 Suporte a Arquivos
*   **Entrada:**
    *   `trimesh` e `meshio` para OBJ, STL, PLY.
    *   `PyVista` nativo para VTP/VTK.
    *   **FBX:** Requer conversão via `Assimp` (se licença permitir) ou conversor externo, pois o SDK é proprietário.
    *   **STP (CAD):** Será tesselado usando `gmsh` ou `pythonocc` para converter NURBS em malha triangular processável.
*   **Saída (Navisworks):**
    *   Exportação de CSV contendo `Object_ID`, `Status (Wet/Dry)`, `Coverage_%`.
    *   Opcionalmente, exportar a malha colorida como FBX ou NWD (via plugin proprietário, se disponível) ou VTP colorido (padrão aberto).

## 5. Interface Lateral e UX
A interface lateral (`SidePanel`) conterá:
1.  **Árvore de Objetos:** Lista hierárquica de meshes e bicos.
2.  **Inspetor de Propriedades:** Ajuste fino de posição (X,Y,Z), rotação e ângulo do cone selecionado.
3.  **Barra de Progresso:** Feedback visual durante o Ray Casting (conectada via `QThread` signals).
4.  **Heatmap Control:** Slider para ajustar a opacidade da visualização de molhado/seco sobre a textura original.
