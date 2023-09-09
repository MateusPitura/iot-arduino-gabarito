/* Máquina corretora de gabarito integrada ao Excel
  Identifica as alternativas marcadas e as corrige com base no gabarito informado via Excel, exibindo a correção nele
  Mateus Enrick Pitura
  26/03/2021 */ 
 
//1. CONTANTES
const int luminosidadeDefinida = 20; // Usada como parâmetro para saber se uma alternativa foi marcada
const int tempoDefinidoCorrecao = 2000; // Tempo total de correção, mínimo recomendado: 2000ms
 
//1.1 Entradas
const byte entradaBotaoParar = 13;
 
//1.1.1 recebe a leitura de luminosidade do LDR pela porta analógica
const int entradaLuminosidadeA = 1; //36
const int entradaLuminosidadeB = 2; //32
const int entradaLuminosidadeC = 3; //34
const int entradaLuminosidadeD = 4; //39
const int entradaLuminosidadeE = 5; //35
 
//1.2 Saídas
const byte saidaLedVermelho = 10; //9
const byte saidaLedVerde = 11; //10
const byte saidaLedAmarelo = 12; //11
 
//1.2.1 ativa a leitura do LDR de uma determinada linha
const byte saidaQuestao1 = 2; //33
const byte saidaQuestao2 = 3; //27
const byte saidaQuestao3 = 4; //26
const byte saidaQuestao4 = 5; //25
const byte saidaQuestao5 = 6; //23
 
//2. DADOS NUMÉRICOS
int quantidadeProva = 0; // Quantidade de provas definida pelo usuário
int notaTotal = 0; // Nota total/final de uma prova corrigida
int repeticaoLoopFor = 0; // Quantidade de repetições de um loop/laço for
int numeroOrdinarioQuestao = 1; // Na correção sem nomes, mostra se é o resultado da 1ª ou 2ª ou 3ª... prova
unsigned long ultimaLeituraMillis = 0; // Armazena o valor do millis
 
//2.1 Verifica se está em branco ou se há 2 ou + respostas por questão
byte verificaResposta1 = 0;
byte verificaResposta2 = 0;
byte verificaResposta3 = 0;
byte verificaResposta4 = 0;
byte verificaResposta5 = 0;
 
//2.2 Armazena a leitura de luminosidade dos LDRs
int luminosidadeLidaA = 0;
int luminosidadeLidaB = 0;
int luminosidadeLidaC = 0;
int luminosidadeLidaD = 0;
int luminosidadeLidaE = 0;
 
//2.3 Armazena o valo/nota de cada questão
int valorQuestao1 = 0;
int valorQuestao2 = 0;
int valorQuestao3 = 0;
int valorQuestao4 = 0;
int valorQuestao5 = 0;

//2.4 Armazena a soma das leituras dos LDRs de cada alternativa
int somaLeituraA = 0;
int somaLeituraB = 0;
int somaLeituraC = 0;
int somaLeituraD = 0;
int somaLeituraE = 0;

//2.5 Armazena o número de repetições do processo de leitura dos LDRs de cada questão
int contaRepeticoes1 = 0;
int contaRepeticoes2 = 0;
int contaRepeticoes3 = 0;
int contaRepeticoes4 = 0;
int contaRepeticoes5 = 0;

//2.6 Armazena a média da leitura do LDRs de cada alternativa
int mediaA = 0;
int mediaB = 0;
int mediaC = 0;
int mediaD = 0;
int mediaE = 0;
 
//3. DADOS BOOLEANOS
bool estadoBotaoParar = LOW;
bool iniciaGabaritoGeral = false; // Inicia a definição do gabarito quando recebe um comando do excel
bool travaGabarito = false; // Impede que uma definição de gabarito seja iniciada enquanto outra está em processamento
bool iniciaCorrecao = false; // Inicia a correção quando recebe um comando do excel
bool habilitaCorrecao = false; // Habilita o processo de correção
bool leituraResposta = false; // Inicia a leitura das respostas/alternativas marcadas
bool insercaoQuantidadeProva = false; // Inicia a inserção da quantidade de provas
bool insercaoValor = false; // Inicia a definição da nota/valor de cada questão
bool habilitaLeituraExcel = true; // Habilita a função que recebe os comandos de controle do excel
 
//3.1 Inicia a definição do gabarito (tanto a letra quanto o valor)
bool iniciaGabarito1 = false;
bool iniciaGabarito2 = false;
bool iniciaGabarito3 = false;
bool iniciaGabarito4 = false;
bool iniciaGabarito5 = false;
 
//3.2 Reinicia a definição do gabarito se for digitado uma letra difente das permitidas
bool habilitaValor1 = true;
bool habilitaValor2 = true;
bool habilitaValor3 = true;
bool habilitaValor4 = true;
bool habilitaValor5 = true;
 
//4. DADOS DO TIPO CARACTER
char caractereExcel; // Lê a letra digitada na função que recebe comandos do excel
char letraResposta; // Lê a letra digitada durante a definição do gabarito
 
//4.1 Armazena a letra digitada durante a definição do gabarito
String respostaCerta1 = "";
String respostaCerta2 = "";
String respostaCerta3 = "";
String respostaCerta4 = "";
String respostaCerta5 = "";
 
//4.2 Armazena a resposta/alternativa marcada
String resposta1 = "";
String resposta2 = "";
String resposta3 = "";
String resposta4 = "";
String resposta5 = "";

//Motor
int pino_passo = 9;
int pino_direcao = 8;
int direcao = 1;
int passos_motor1 = 11000;
int passos_motor2 = 7400;
int passos_motor3 = 5000;
 
void setup() {
  Serial.begin(9600);
  pinMode(saidaQuestao1, OUTPUT);
  pinMode(saidaQuestao2, OUTPUT);
  pinMode(saidaQuestao3, OUTPUT);
  pinMode(saidaQuestao4, OUTPUT);
  pinMode(saidaQuestao5, OUTPUT);
  pinMode(saidaLedVermelho, OUTPUT);
  pinMode(entradaBotaoParar, INPUT);
  pinMode(pino_passo, OUTPUT); //Motor
  pinMode(pino_direcao, OUTPUT); //Motor

  //Luzes

  digitalWrite(saidaLedVermelho, HIGH);
  delay(250);
  digitalWrite(saidaLedVermelho, LOW);
  digitalWrite(saidaLedAmarelo, HIGH);
  delay(250);
  digitalWrite(saidaLedAmarelo, LOW);
  digitalWrite(saidaLedVerde, HIGH);
  delay(250);
  digitalWrite(saidaLedVerde, LOW);

  digitalWrite(saidaLedVermelho, HIGH);
  delay(250);
  digitalWrite(saidaLedAmarelo, HIGH);
  delay(250);
  digitalWrite(saidaLedVerde, HIGH);
  delay(500);

  digitalWrite(saidaLedVerde, LOW);
  delay(500);
  digitalWrite(saidaLedAmarelo, LOW);
  delay(500);
  digitalWrite(saidaLedVermelho, LOW);
}

// FUNÇÃO QUE LIMPA O BUFFER (armazenamento temporário de dados) DA SERIAL
void limparBuffer() {
  while (Serial.available() > 0) {
    Serial.read();
  }
}

// FUNÇÃO QUE ZERA OS VALORES DAS VARIÁVEIS USADAS PARA FAZER A MÉDIA DAS LEITURAS DOS LDRs
void zeraMedia(){
  somaLeituraA = 0;
  somaLeituraB = 0;
  somaLeituraC = 0;
  somaLeituraD = 0;
  somaLeituraE = 0;
  mediaA = 0;
  mediaB = 0;
  mediaC = 0;
  mediaD = 0;
  mediaE = 0;
}

// FUNÇÃO COM OS COMANDOS QUE PARA QUALQUER PROCESSO QUE ESTEJA ACONTECENDO
void comandosParar(){
  limparBuffer();
  zeraMedia();
  digitalWrite(saidaLedVerde, LOW);
  digitalWrite(saidaLedAmarelo, LOW);
  iniciaGabarito1 = false;
  iniciaGabarito2 = false;
  iniciaGabarito3 = false;
  iniciaGabarito4 = false;
  iniciaGabarito5 = false;
  habilitaValor1 = false; 
  habilitaValor2 = false;
  habilitaValor3 = false;
  habilitaValor4 = false;
  habilitaValor5 = false;
  leituraResposta = false;
  travaGabarito = false;
  habilitaLeituraExcel = true;
  numeroOrdinarioQuestao = 0;
  repeticaoLoopFor = quantidadeProva * 2; // Impede que a inserção de nomes aconteça
  verificaResposta5 = 10; // Impede que a correção seja exibida quando a função parar for chamada
}
 
// FUNÇÃO QUE CHAMA OS COMANDOS DA FUNÇÃO PARAR
void parar() {
  estadoBotaoParar = digitalRead(entradaBotaoParar);
  if (estadoBotaoParar == HIGH) {
    comandosParar();
  }
}
 
// TRATAMENTO DE ERRO CASO SEJA INSERIDO UMA LETRA DIFERENTE DAS ACEITAS
void erroLetraGabarito(){
  if (letraResposta != 'A' && letraResposta != 'B' && letraResposta != 'C' && letraResposta != 'D' && letraResposta != 'E' && 
      letraResposta != 'a' && letraResposta != 'b' && letraResposta != 'c' && letraResposta != 'd' && letraResposta != 'e') { 
    comandosParar();
    digitalWrite(saidaLedVermelho, HIGH);
    while(1 == 1){} // Entra em um loop infinito no qual a única forma de sair dele é resetando o arduino
  }
}

// TRATAMENTO DE ERRO CASO SEJA INSERIDO UM VALOR DIFERENTE DO ACEITO
void erroValorGabarito(int valorQuestoes){ // Essa variável é o parâmetro que irá receber o valor das questões
  if (valorQuestoes <= 0) {
    insercaoValor = false;
    comandosParar();
    digitalWrite(saidaLedVermelho, HIGH);
    while(1 == 1){}
  }
} 

// FUNÇÃO QUE LÊ E FAZ A MÉDIA DA LUMINOSIDADE DOS LDRs
void lerLuminosidade(int contaRepeticoes) { // Essa variável é o parâmetro que recebe o número de repetições
  luminosidadeLidaA = analogRead(entradaLuminosidadeA);
  somaLeituraA += luminosidadeLidaA;
  mediaA = somaLeituraA / contaRepeticoes;
  luminosidadeLidaB = analogRead(entradaLuminosidadeB);
  somaLeituraB += luminosidadeLidaB;
  mediaB = somaLeituraB / contaRepeticoes;
  luminosidadeLidaC = analogRead(entradaLuminosidadeC);
  somaLeituraC += luminosidadeLidaC;
  mediaC = somaLeituraC / contaRepeticoes;
  luminosidadeLidaD = analogRead(entradaLuminosidadeD);
  somaLeituraD += luminosidadeLidaD;
  mediaD = somaLeituraD / contaRepeticoes;
  luminosidadeLidaE = analogRead(entradaLuminosidadeE);
  somaLeituraE += luminosidadeLidaE;
  mediaE = somaLeituraE / contaRepeticoes;
}
 
void loop() {
  parar();

  // LEITURA DE COMANDOS DO EXCEL
  while (habilitaLeituraExcel == true) {
    if ((millis() - ultimaLeituraMillis) >= 5000){ // Essa leitura do millis acontece na parte que acende o led verde depois de corrigir todas as provas
      digitalWrite(saidaLedVerde, LOW);
    }
    parar();
    caractereExcel = Serial.read();
    if (caractereExcel == 'G'){
      delay(10);
      iniciaGabaritoGeral = true;
      habilitaLeituraExcel = false;
    } 
    if (caractereExcel == 'R'){
      delay(10);
      iniciaCorrecao = true;
      habilitaLeituraExcel = false;
    }
  }
 
  // PROCESSO DE DEFINIÇÃO DE GABARITOS E NOTAS
  if (iniciaGabaritoGeral == true && travaGabarito == false) {
    iniciaGabaritoGeral = false;
    limparBuffer();
    digitalWrite(saidaLedVerde, LOW);
    digitalWrite(saidaLedAmarelo, HIGH);
    iniciaGabarito1 = true;
    travaGabarito = true;
    habilitaCorrecao = false;
    numeroOrdinarioQuestao = 1;
  }
 
  // Inicia gabarito 1
  if (iniciaGabarito1 == true) {
    respostaCerta1 = "";
    if (Serial.available() > 0) {
      letraResposta = Serial.read();
      respostaCerta1 += letraResposta;
      habilitaValor1 = true;      
      erroLetraGabarito();
      if (habilitaValor1 == true) { // Inicia nota
        limparBuffer();
        insercaoValor = true;
        valorQuestao1 = 0;
        while (insercaoValor == true && estadoBotaoParar == LOW) { 
          parar();
          if (Serial.available() > 0) {
            Serial.setTimeout(10); // Torna a leitura do número inteiro mais rápida
            valorQuestao1 = Serial.parseInt(); // Lê um número inteiro digitado na entrada serial
            if (valorQuestao1 > 0) {
              limparBuffer();
              iniciaGabarito2 = true;
              iniciaGabarito1 = false;
              insercaoValor = false;
            }
            erroValorGabarito(valorQuestao1);
            delay(10);
          }
        }
      }
    }
  }
 
  // Inicia gabarito 2
  if (iniciaGabarito2 == true) {
    respostaCerta2 = "";
    if (Serial.available() > 0) {
      letraResposta = Serial.read();
      respostaCerta2 += letraResposta;
      habilitaValor2 = true;
      erroLetraGabarito();
      if (habilitaValor2 == true) {
        limparBuffer();
        insercaoValor = true;
        valorQuestao2 = 0;
        while (insercaoValor == true && estadoBotaoParar == LOW) {
          parar();
          if (Serial.available() > 0) {
            Serial.setTimeout(10);
            valorQuestao2 = Serial.parseInt();
            if (valorQuestao2 > 0) {
              limparBuffer();
              iniciaGabarito3 = true;
              iniciaGabarito2 = false;
              insercaoValor = false;
            }
            erroValorGabarito(valorQuestao2);
            delay(10);
          }
        }
      }
    }
  }
 
  // Inicia gabarito 3
  if (iniciaGabarito3 == true) {
    respostaCerta3 = "";
    if (Serial.available() > 0) {
      letraResposta = Serial.read();
      respostaCerta3 += letraResposta;
      habilitaValor3 = true;
      erroLetraGabarito();
      if (habilitaValor3 == true) {
        limparBuffer();
        insercaoValor = true;
        valorQuestao3 = 0;
        while (insercaoValor == true && estadoBotaoParar == LOW) {
          parar();
          if (Serial.available() > 0) {
            Serial.setTimeout(10);
            valorQuestao3 = Serial.parseInt();
            if (valorQuestao3 > 0) {
              limparBuffer();
              iniciaGabarito4 = true;
              iniciaGabarito3 = false;
              insercaoValor = false;
            }
            erroValorGabarito(valorQuestao3);
            delay(10);
          }
        }
      }
    }
  }
 
  // Inicia gabarito 4
  if (iniciaGabarito4 == true) {
    respostaCerta4 = "";
    if (Serial.available() > 0) {
      letraResposta = Serial.read();
      respostaCerta4 += letraResposta;
      habilitaValor4 = true;
      erroLetraGabarito();
      if (habilitaValor4 == true) {
        limparBuffer();
        insercaoValor = true;
        valorQuestao4 = 0;
        while (insercaoValor == true && estadoBotaoParar == LOW) {
          parar();
          if (Serial.available() > 0) {
            Serial.setTimeout(10);
            valorQuestao4 = Serial.parseInt();
            if (valorQuestao4 > 0) {
              limparBuffer();
              iniciaGabarito5 = true;
              iniciaGabarito4 = false;
              insercaoValor = false;
            }
            erroValorGabarito(valorQuestao4);
            delay(10);
          }
        }
      }
    }
  }
 
  // Inicia gabarito 5
  if (iniciaGabarito5 == true) {
    respostaCerta5 = "";
    if (Serial.available() > 0) {
      letraResposta = Serial.read();
      respostaCerta5 += letraResposta;
      habilitaValor5 = true;
      erroLetraGabarito();
      if (habilitaValor5 == true) {
        limparBuffer();
        insercaoValor = true;
        valorQuestao5 = 0;
        while (insercaoValor == true && estadoBotaoParar == LOW) {
          parar();
          if (Serial.available() > 0) {
            Serial.setTimeout(10);
            valorQuestao5 = Serial.parseInt();
            if (valorQuestao5 > 0) {
              limparBuffer();
              digitalWrite(saidaLedAmarelo, LOW);
              habilitaCorrecao = true;
              iniciaGabarito5 = false;
              travaGabarito = false;
              insercaoValor = false;
              habilitaLeituraExcel = true;
            }
            erroValorGabarito(valorQuestao5);
            delay(10);
          }
        }
      }
    }
  }
 
  // PROCESSO DE INSERÇÃO DE INFORMAÇÕES NECESSÁRIAS PARA A CORREÇÃO
  if (iniciaCorrecao == true && habilitaCorrecao == true) { 
    iniciaCorrecao = false;
    limparBuffer();
    digitalWrite(saidaLedAmarelo, HIGH);
    digitalWrite(saidaLedVerde, LOW);
    leituraResposta = false;
    insercaoQuantidadeProva = false;
    quantidadeProva = 0;
    repeticaoLoopFor = 1;
    numeroOrdinarioQuestao = 1;
    
    // Definição da quantidade de provas que serão corrigidas
    while (insercaoQuantidadeProva == false && estadoBotaoParar == LOW) {
      parar();
      if (Serial.available() > 0) {
        Serial.setTimeout(10);
        quantidadeProva = Serial.parseInt();
        if (quantidadeProva > 0) {
          insercaoQuantidadeProva = true;
          Serial.println("");
          Serial.print(respostaCerta1 + " - ");
          Serial.println(valorQuestao1);
          Serial.print(respostaCerta2 + " - ");
          Serial.println(valorQuestao2);
          Serial.print(respostaCerta3 + " - ");
          Serial.println(valorQuestao3);
          Serial.print(respostaCerta4 + " - ");
          Serial.println(valorQuestao4);
          Serial.print(respostaCerta5 + " - ");
          Serial.println(valorQuestao5);
          Serial.println(valorQuestao1 + valorQuestao2 + valorQuestao3 + valorQuestao4 + valorQuestao5);
          leituraResposta = true; 
          digitalWrite(saidaLedAmarelo, LOW); 
        }
        if (quantidadeProva <= 0) {
          insercaoQuantidadeProva = true;
          comandosParar();
          digitalWrite(saidaLedVermelho, HIGH);
          while(1 == 1){}
        }
        delay(10);
      }
    }

    // PROCESSO DE CORREÇÃO
 
    // Processo que lê a luminosidade dos LDRs
    if (leituraResposta == true && estadoBotaoParar == LOW) {
      for (repeticaoLoopFor = 1; repeticaoLoopFor <= quantidadeProva; repeticaoLoopFor++ && estadoBotaoParar == LOW) {
        ultimaLeituraMillis = millis();
        digitalWrite(saidaLedVerde, LOW);
        notaTotal = 0;
        contaRepeticoes1 = 0;
        contaRepeticoes2 = 0;
        contaRepeticoes3 = 0;
        contaRepeticoes4 = 0;
        contaRepeticoes5 = 0;

        // Motor inicio 
        
          direcao = 0;
          digitalWrite(pino_direcao, direcao);
          for (int p=0 ; p < passos_motor1; p++){
            parar();
            digitalWrite(pino_passo, 1);
            delay(1);
            digitalWrite(pino_passo, 0);
            delay(1);
          }
          
          delay(1000);
          
          direcao = 1;
          digitalWrite(pino_direcao, direcao);
          for (int p=0 ; p < passos_motor2; p++){
            parar();
            digitalWrite(pino_passo, 1);
            delay(1);
            digitalWrite(pino_passo, 0);
            delay(1);
          }
        
          delay(1000);
        // Motor fim

        ultimaLeituraMillis = millis();
        
        while ((millis() - ultimaLeituraMillis) <= tempoDefinidoCorrecao && estadoBotaoParar == LOW && numeroOrdinarioQuestao != 0) { // Executa o processo abaixo durante o tempo total de correção definido
          parar();
          if ((millis() - ultimaLeituraMillis) > 0.01 && (millis() - ultimaLeituraMillis) <= (tempoDefinidoCorrecao * 0.19) && contaRepeticoes1 < 5) { // Define uma faixa de tempo na qual será executado apenas o processo abaixo
            digitalWrite(saidaQuestao1, HIGH);
            digitalWrite(saidaLedAmarelo, LOW);
            verificaResposta1 = 0;
            lerLuminosidade(++contaRepeticoes1);
            if (contaRepeticoes1 == 5){
              if (mediaA <= luminosidadeDefinida) {
                resposta1 = "A";
                verificaResposta1++;
              }
              if (mediaB <= luminosidadeDefinida) {
                resposta1 = "B";
                verificaResposta1++;
              }
              if (mediaC <= luminosidadeDefinida) {
                resposta1 = "C";
                verificaResposta1++;
              }
              if (mediaD <= luminosidadeDefinida) {
                resposta1 = "D";
                verificaResposta1++;
              }
              if (mediaE <= luminosidadeDefinida) {
                resposta1 = "E";
                verificaResposta1++;
              }
              digitalWrite(saidaQuestao1, LOW);
              zeraMedia();
            }
          }
 
          if ((millis() - ultimaLeituraMillis) > (tempoDefinidoCorrecao * 0.21) && (millis() - ultimaLeituraMillis) <= (tempoDefinidoCorrecao * 0.39) && contaRepeticoes2 < 5) {
            digitalWrite(saidaQuestao2, HIGH);
            digitalWrite(saidaLedAmarelo, HIGH);
            verificaResposta2 = 0;
            lerLuminosidade(++contaRepeticoes2);
            if (contaRepeticoes2 == 5){
              if (luminosidadeLidaA <= luminosidadeDefinida) {
                resposta2 = "A";
                verificaResposta2++;
              }
              if (luminosidadeLidaB <= luminosidadeDefinida) {
                resposta2 = "B";
                verificaResposta2++;
              }
              if (luminosidadeLidaC <= luminosidadeDefinida) {
                resposta2 = "C";
                verificaResposta2++;
              }
              if (luminosidadeLidaD <= luminosidadeDefinida) {
                resposta2 = "D";
                verificaResposta2++;
              }
              if (luminosidadeLidaE <= luminosidadeDefinida) {
                resposta2 = "E";
                verificaResposta2++;
              }
              digitalWrite(saidaQuestao2, LOW);
              zeraMedia();
            }
          }
 
          if ((millis() - ultimaLeituraMillis) > (tempoDefinidoCorrecao * 0.41) && (millis() - ultimaLeituraMillis) <= (tempoDefinidoCorrecao * 0.59) && contaRepeticoes3 < 5) {
            digitalWrite(saidaQuestao3, HIGH);
            digitalWrite(saidaLedAmarelo, LOW);
            verificaResposta3 = 0;
            lerLuminosidade(++contaRepeticoes3);
            if (contaRepeticoes3 == 5){
              if (luminosidadeLidaA <= luminosidadeDefinida) {
                resposta3 = "A";
                verificaResposta3++;
              }
              if (luminosidadeLidaB <= luminosidadeDefinida) {
                resposta3 = "B";
                verificaResposta3++;
              }
              if (luminosidadeLidaC <= luminosidadeDefinida) {
                resposta3 = "C";
                verificaResposta3++;
              }
              if (luminosidadeLidaD <= luminosidadeDefinida) {
                resposta3 = "D";
                verificaResposta3++;
              }
              if (luminosidadeLidaE <= luminosidadeDefinida) {
                resposta3 = "E";
                verificaResposta3++;
              }
              digitalWrite(saidaQuestao3, LOW);
              zeraMedia();
            }
          }
 
          if ((millis() - ultimaLeituraMillis) > (tempoDefinidoCorrecao * 0.61) && (millis() - ultimaLeituraMillis) <= (tempoDefinidoCorrecao * 0.79) && contaRepeticoes4 < 5) {
            digitalWrite(saidaQuestao4, HIGH);
            digitalWrite(saidaLedAmarelo, HIGH);
            verificaResposta4 = 0;
            lerLuminosidade(++contaRepeticoes4);
            if (contaRepeticoes4 == 5){
              if (luminosidadeLidaA <= luminosidadeDefinida) {
                resposta4 = "A";
                verificaResposta4++;
              }
              if (luminosidadeLidaB <= luminosidadeDefinida) {
                resposta4 = "B";
                verificaResposta4++;
              }
              if (luminosidadeLidaC <= luminosidadeDefinida) {
                resposta4 = "C";
                verificaResposta4++;
              }
              if (luminosidadeLidaD <= luminosidadeDefinida) {
                resposta4 = "D";
                verificaResposta4++;
              }
              if (luminosidadeLidaE <= luminosidadeDefinida) {
                resposta4 = "E";
                verificaResposta4++;
              }
              digitalWrite(saidaQuestao4, LOW);
              zeraMedia();
            }
          }
 
          if ((millis() - ultimaLeituraMillis) > (tempoDefinidoCorrecao * 0.81) && (millis() - ultimaLeituraMillis) <= (tempoDefinidoCorrecao * 0.99) && contaRepeticoes5 < 5) {
            digitalWrite(saidaQuestao5, HIGH);
            digitalWrite(saidaLedAmarelo, LOW);
            verificaResposta5 = 0;
            lerLuminosidade(++contaRepeticoes5);
            if (contaRepeticoes5 == 5){
              if (luminosidadeLidaA <= luminosidadeDefinida) {
                resposta5 = "A";
                verificaResposta5++;
              }
              if (luminosidadeLidaB <= luminosidadeDefinida) {
                resposta5 = "B";
                verificaResposta5++;
              }
              if (luminosidadeLidaC <= luminosidadeDefinida) {
                resposta5 = "C";
                verificaResposta5++;
              }
              if (luminosidadeLidaD <= luminosidadeDefinida) {
                resposta5 = "D";
                verificaResposta5++;
              }
              if (luminosidadeLidaE <= luminosidadeDefinida) {
                resposta5 = "E";
                verificaResposta5++;
              }
              digitalWrite(saidaQuestao5, LOW);
              zeraMedia();
            }
          }
          delay(10);
        }

        // Exibi o número de cada prova se a resposta para a pergunta foi "nao"
        if (numeroOrdinarioQuestao == 0) {}
        else {
          Serial.print(numeroOrdinarioQuestao);
          Serial.println("' PROVA");
          numeroOrdinarioQuestao++;
        }

        // Processo de correção (comparação da alterantiva marcada com a resposta certa)
        if (resposta1.equalsIgnoreCase(respostaCerta1) && verificaResposta1 == 1) { // Compara se a reposta é igual a resposta certa, sem diferenciar maiúscula de minúscula
          Serial.println("CORRETA");
          notaTotal += valorQuestao1;
        }
        else {
          if (verificaResposta1 > 1) {
            if (verificaResposta5 == 10) {} // Acontece quando o botão parar é precionado, não executa nenhum código, mas é importante para nenhuma mensagem aparecer
            else Serial.println("ANULADA (2 ou + alternativas)");
          }
          if (verificaResposta1 == 0) {
            Serial.println("ANULADA (em branco)");
          }
          if (verificaResposta1 == 1) {
            Serial.println("ERRADA");
          }
        }
 
        if (resposta2.equalsIgnoreCase(respostaCerta2) && verificaResposta2 == 1) {
          Serial.println("CORRETA");
          notaTotal += valorQuestao2;
        }
        else {
          if (verificaResposta2 > 1) {
            if (verificaResposta5 == 10) {}
            else Serial.println("ANULADA (2 ou + alternativas)");
          }
          if (verificaResposta2 == 0) {
            Serial.println("ANULADA (em branco)");
          }
          if (verificaResposta2 == 1) {
            Serial.println("ERRADA");
          }
        }
 
        if (resposta3.equalsIgnoreCase(respostaCerta3) && verificaResposta3 == 1) {
          Serial.println("CORRETA");
          notaTotal += valorQuestao3;
        }
        else {
          if (verificaResposta3 > 1) {
            if (verificaResposta5 == 10) {}
            else Serial.println("ANULADA (2 ou + alternativas)");
          }
          if (verificaResposta3 == 0) {
            Serial.println("ANULADA (em branco)");
          }
          if (verificaResposta3 == 1) {
            Serial.println("ERRADA");
          }
        }
 
        if (resposta4.equalsIgnoreCase(respostaCerta4) && verificaResposta4 == 1) {
          Serial.println("CORRETA");
          notaTotal += valorQuestao4;
        }
        else {
          if (verificaResposta4 > 1) {
            if (verificaResposta5 == 10) {}
            else Serial.println("ANULADA (2 ou + alternativas)");
          }
          if (verificaResposta4 == 0) {
            Serial.println("ANULADA (em branco)");
          }
          if (verificaResposta4 == 1) {
            Serial.println("ERRADA");
          }
        }
 
        if (resposta5.equalsIgnoreCase(respostaCerta5) && verificaResposta5 == 1) {
          Serial.println("CORRETA");
          notaTotal += valorQuestao5;
        }
        else {
          if (verificaResposta5 > 1) {
            if (verificaResposta5 == 10) {}
            else Serial.println("ANULADA (2 ou + alternativas)"); 
          }
          if (verificaResposta5 == 0) {
            Serial.println("ANULADA (em branco)");
          }
          if (verificaResposta5 == 1) {
            Serial.println("ERRADA");
          }
        }
 
        // Processo que exibe a nota total/final
        if (verificaResposta5 != 10) {
          Serial.println(notaTotal);
        }
 
        // Processo que acende o LED verde depois de corrigir a última prova
        if (repeticaoLoopFor == quantidadeProva) {
          digitalWrite(saidaLedVerde, HIGH);
          habilitaLeituraExcel = true;
          ultimaLeituraMillis = millis();
        }

        // Motor inicio
        
          direcao = 1;
          digitalWrite(pino_direcao, direcao);
          for (int p=0 ; p < passos_motor3; p++){
            parar();
            digitalWrite(pino_passo, 1);
            delay(1);
            digitalWrite(pino_passo, 0);
            delay(1);
          }
          delay(1000); 
        // Motor fim
        
      }
    }
  }
  delay(10);
}