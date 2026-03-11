# Geospatial Distance Optimizer & Proximity Engine (VBA)

## 📌 Sobre o Projeto

Este é um projeto avançado de **Inteligência Geográfica** desenvolvido inteiramente em VBA para Excel. A ferramenta é capaz de geolocalizar endereços, calcular distâncias entre múltiplos pontos considerando a curvatura da Terra e gerar matrizes de proximidade para apoio à decisão logística e estratégica.

O sistema não apenas calcula distâncias lineares, mas processa uma rede de pontos para determinar o ordenamento de proximidade e filtragem por raio de ação (Geofencing).

### 💡 Complexidade e Diferenciais Técnicos

Diferente de ferramentas convencionais, este projeto resolve problemas complexos de engenharia de dados:
* **Geocodificação Dinâmica:** Integração com a API *Nominatim (OpenStreetMap)* via requisições HTTP para converter endereços em coordenadas de Latitude e Longitude em tempo real.
* **Cálculo Esférico (Haversine):** Implementação da Fórmula de Haversine para garantir precisão nos cálculos de distância, considerando o raio médio da Terra ($6.371$ km).
* **Processamento de Matriz $N \times N$:** Algoritmo que gera uma matriz completa de distâncias entre todos os pontos da rede, permitindo análises de "Clusterização".
* **Ordenação de Dados (Bubble Sort):** Implementação de lógica de ordenação em memória (Arrays) para entregar resultados instantâneos de "mais próximo para mais longe".

---

## 🏗️ Funcionalidades do Sistema

### 1. Sincronização Geográfica (API Integration)
O módulo realiza o tratamento de strings e consulta servidores externos para enriquecer a base de dados com metadados geográficos.

### 2. Matriz de Distâncias Centralizada
Gera uma visão cruzada de toda a rede de filiais, fundamental para entender a sobreposição de áreas de atuação.

![matriz](https://github.com/user-attachments/assets/e33d7d33-b0a1-44ea-8e83-ea3510ab68f9)

*Figura 1: Matriz de Distâncias Geoprocessada.*

### 3. Painel de Análise e Proximidade
Interface de usuário onde é possível:
* Selecionar um ponto de referência.
* Obter a lista completa de filiais ordenadas da mais próxima para a mais distante.
* Aplicar filtros de raio (ex: "Quais filiais estão a menos de 15km deste ponto?").

![distancias](https://github.com/user-attachments/assets/ba1877a5-eeaa-4134-87cd-933be75dfb70)

*Figura 2: Painel de Controle e Filtros de Raio de Proximidade.*

---

## 🛠️ Fundamentação Matemática

Para o cálculo de distância entre dois pontos ($P_1$ e $P_2$) na superfície da Terra, utilizamos a formulação de Haversine:

$$a = \sin^2\left(\frac{\Delta\phi}{2}\right) + \cos(\phi_1) \cdot \cos(\phi_2) \cdot \sin^2\left(\frac{\Delta\lambda}{2}\right)$$
$$c = 2 \cdot \operatorname{atan2}(\sqrt{a}, \sqrt{1-a})$$
$$d = R \cdot c$$

Onde:
* $\phi$ = Latitude (em radianos)
* $\lambda$ = Longitude (em radianos)
* $R$ = Raio da Terra ($6371$ km)

---

## 🛡️ Disclaimer e Privacidade (Dados Anonimizados)

* **Dados de Teste:** Todos os endereços e coordenadas utilizados neste repositório foram gerados aleatoriamente ou extraídos de bases públicas de endereços reais de forma desconexa.
* **Privacidade:** Nomes de filiais, IDs de filiais e valores estratégicos foram removidos ou substituídos para proteger a confidencialidade do negócio original. O projeto está **pronto para uso**, bastando inserir novas chaves de endereços.
