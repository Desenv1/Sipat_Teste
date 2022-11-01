// Determinando as colunas da planilha "Dados dos responsáveis" como globais para facilitar alterações.
var colOrgao = 1;
var colSetor = 2;
var colAlias = 3;
var colSala = 4;
var colNome = 5;
var colEmail = 6;
var colSincroniza = 11;

SORT_ORDER = [
	{column: 8, ascending: true},  // 1 = column number, sorting by ascending order
	{column: 1, ascending: true}, // 2 = column number, sort by ascending order 
	{column: 2, ascending: true},
	{column: 3, ascending: true},
	{column: 4, ascending: true},
	{column: 5, ascending: true},
	{column: 6, ascending: true},
	{column: 7, ascending: true},
	{column: 8, ascending: true},
	{column: 10, ascending: true},
	]; // Para organização dos patrimônios na planilha mestre.
	
	SORT_ORDER2 = [
	{column: 4, ascending: true},  // 1 = column number, sorting by descending order
	{column: 1, ascending: true}, // 2 = column number, sort by ascending order 
	{column: 2, ascending: true},
	{column: 3, ascending: true},
	{column: 5, ascending: true},
	{column: 6, ascending: true},
	]; // Para organização dos cadastros na planilha de dados dos responsáveis.
	
