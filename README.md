# Otimizador de Distância Geográfica e Motor de Proximidade (VBA)

## 📌 Sobre o Projeto

Este é um projeto de **Inteligência Geográfica** desenvolvido inteiramente em VBA para Excel. A ferramenta é capaz de geolocalizar endereços, calcular distâncias entre múltiplos pontos considerando a curvatura da Terra e gerar matrizes de proximidade para apoio à decisão logística e estratégica.

O sistema não apenas calcula distâncias lineares, mas processa uma rede de pontos para determinar o ordenamento de proximidade e filtragem por raio de ação (*Geofencing*).

### 💡 Complexidade e Diferenciais Técnicos

Este projeto resolve problemas complexos de engenharia de dados dentro do ambiente Office:
* **Geocodificação Dinâmica:** Integração com a API *Nominatim* via requisições HTTP para converter endereços em coordenadas de Latitude e Longitude em tempo real.
* **Cálculo Esférico (Haversine):** Implementação da Fórmula de Haversine para garantir precisão nos cálculos de distância, considerando o raio médio da Terra de 6.371 km.
* **Processamento de Matriz $N \times N$:** Algoritmo que gera uma matriz completa de distâncias entre todos os pontos da rede.
* **Ordenação de Dados (Bubble Sort):** Implementação de lógica de ordenação em memória para entregar resultados instantâneos de "mais próximo para mais longe".

---

## 🏗️ Funcionalidades do Sistema

### 1. Sincronização Geográfica (API Integration)
O módulo realiza o tratamento de strings e consulta servidores externos para enriquecer a base de dados com metadados geográficos automaticamente.

### 2. Matriz de Distâncias Centralizada
Gera uma visão cruzada de toda a rede de filiais, fundamental para entender a sobreposição de áreas de atuação.

![matriz](https://github.com/user-attachments/assets/0421cefe-86f0-41a2-b03e-705b8b78d4a6)

*Figura 1: Matriz de Distâncias Geoprocessada.*

### 3. Painel de Análise e Proximidade
Interface de usuário onde é possível filtrar dados de forma inteligente:
* Selecionar um ponto de referência.
* Obter a lista completa de filiais ordenadas por distância.
* Aplicar filtros de raio (ex: filiais em um raio de 15km).

![distancias](https://github.com/user-attachments/assets/1cdd6d72-2ae8-4ea1-a132-ec3a7e598ee3)

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
* $R$ = Raio da Terra (6.371 km)

---

## 🛡️ Disclaimer e Privacidade (Dados Anonimizados)

* **Dados de Teste:** Todos os endereços e coordenadas utilizados foram gerados aleatoriamente ou extraídos de bases públicas de forma desconexa.
* **Pronto para Uso:** O projeto está estruturado para ser escalável, bastando inserir novos endereços para que a automação processe as coordenadas.
