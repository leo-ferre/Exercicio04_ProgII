# Sistema de Controle de Estoque
## Exercício 04 - Programação II

Sistema para gerenciar estoque usando Python e Excel (biblioteca openpyxl).
- Feito por:
- Heloi Vecchi Sgarbi
- Kaua Schiavolin Monteiro
- Leonardo Ferreira
---

## O que faz

- Cadastra produtos (código, nome, categoria, quantidade, preço)
- Registra entradas e saídas
- Gera relatórios automáticos com fórmulas Excel
- Avisa quando o estoque está baixo

---

## Como usar

### Instalar biblioteca

```bash
python3 -m pip install openpyxl
```

### Executar

```bash
python3 main.py
```

---

## Arquivo Excel

O programa cria `controle_estoque.xlsx` com 3 abas:

**Produtos** - cadastro de itens
- Fórmulas: Valor Total (quantidade × preço) e Status (BAIXO ou OK)

**Movimentações** - histórico de entradas/saídas
- Registra automaticamente data/hora, código, tipo e quantidade

**Relatórios** - estatísticas gerais

---

## Menu do Sistema

```
1. Adicionar novo produto
2. Registrar entrada de estoque
3. Registrar saída de estoque
4. Listar todos os produtos
5. Atualizar relatórios
6. Sair
```

## Exemplos de teste

### Adicionar produtos

```
Código: P001
Nome: Mouse Gamer
Categoria: Periféricos
Quantidade: 50
Estoque Mínimo: 10
Preço: 150.00
```

```
Código: P002
Nome: Teclado Mecânico
Categoria: Periféricos
Quantidade: 5
Estoque Mínimo: 10
Preço: 350.00
```
(Este vai aparecer como estoque BAIXO)

### Registrar entrada

```
Código: P002
Quantidade: 20
```

### Registrar saída

```
Código: P001
Quantidade: 10
```

---

## Tecnologias usadas

- Python 3
- openpyxl (para manipular Excel)
- datetime (para data/hora)
- os (para verificar arquivos)

---

Exercício 04 - Programação II
Novembro/2025
