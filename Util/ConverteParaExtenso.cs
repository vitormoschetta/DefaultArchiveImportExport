using System;

namespace DefaultArchiveImportExport.Util
{
    public static class ConverteParaExtenso
    {
        public static string NumeroParaExtenso(int valor)
        {
            var separador = " e ";

            if(valor < 20)
            {	
                return RetornaInteiroExtenso(valor);
            }
            
            if(valor > 19)
            {	
                int len = valor.ToString().Length;
            
                if(len == 2)
                {
                    int ValorPrimario = int.Parse(valor.ToString().Substring(0,1));
                    int ValorSecundario = int.Parse(valor.ToString().Substring(1,1));
                    ValorPrimario = ValorPrimario * 10;
                    return RetornaInteiroExtenso(ValorPrimario)+(ValorSecundario>0 ? separador+RetornaInteiroExtenso(ValorSecundario):"");
                }
                else if(len == 3)
                {
                    int ValorPrimario = int.Parse(valor.ToString().Substring(0,1));
                    int ValorSecundario = int.Parse(valor.ToString().Substring(1,1));
                    int ValorTerciario = int.Parse(valor.ToString().Substring(2,1));
                    
                    ValorPrimario = ValorPrimario * 100;
                    ValorSecundario = ValorSecundario * 10;
                    
                    
                    return RetornaInteiroExtenso(ValorPrimario)
                                    +(ValorSecundario > 0  ? separador+RetornaInteiroExtenso(ValorSecundario):"")
                                    +(ValorTerciario > 0 ?separador+RetornaInteiroExtenso(ValorTerciario):"");
                }

                
            }

            return "Valor inválido";
        }


        public static string RetornaInteiroExtenso(int identificador)
        {
            switch (identificador)
            {
                case 1: return "Um";
                case 2: return "Dois";
                case 3: return "Três";
                case 4: return "Quatro";
                case 5: return "Cinco";
                case 6: return "Seis";
                case 7: return "Sete";
                case 8: return "Oito";
                case 9: return "Nove";
                
                case 10: return"Dez";
                case 11: return "Onze";
                case 12: return "Doze";
                case 13: return "Treze";
                case 14: return "Quatorze";
                case 15: return "Quinze";
                case 16: return "Dezesseis";
                case 17: return "Dezessete";
                case 18: return "Dezoito";
                case 19: return "Dezenove";
                
                case 20: return"Vinte";
                case 30: return"Trinta";
                case 40: return"Quarenta";
                case 50: return"Cinquenta";
                case 60: return"Sessenta";
                case 70: return"Setenta";
                case 80: return"Oitenta"; 
                case 90: return"Noventa";  
                
                case 100:return "Cem";         
                case 200:return "Duzentos";    
                case 300:return "Trezentos";   
                case 400:return "Quatrocentos";
                case 500:return "Quinhentos";  
                case 600:return "Seicentos";   
                case 700:return "Setecentos";  
                case 800:return "Oitocentos";  
                case 900:return "Novecentos";  
                
                default:return "Valor inválido";
            }
        
        }


        public static string ValorParaExtenso(decimal valor)
        {            
            string milhar = string.Empty;
            string centena = string.Empty;
            string dezena= string.Empty;
            string unidade = string.Empty;
            string extenso = string.Empty;
            string zero = "0000";

            string numero = valor.ToString();
            numero = numero.PadLeft(4, '0');

            if (numero == zero)
            {
                return "Zero";
            }


            // ESCREVER POR EXTENSO MILHAR
            if (numero[0] != '0')
            {
                switch (numero[0])
                {
                    case '1': milhar = "Um Mil"; break;
                    case '2': milhar = "Dois Mil"; break;
                    case '3': milhar = "Três Mil"; break;
                    case '4': milhar = "Quatro Mil"; break;
                    case '5': milhar = "Cinco Mil"; break;
                    case '6': milhar = "Seis Mil"; break;
                    case '7': milhar = "Sete Mil"; break;
                    case '8': milhar = "Oito Mil"; break;
                    case '9': milhar = "Nove Mil"; break;

                }
            }


            // ESCREVER POR EXTENSO CENTENA          
            if (numero[1] != '0')
            {              
                switch (numero[1])
                {
                    case '1': centena = "Cento "; break;
                    case '2': centena = "Duzentos "; break;
                    case '3': centena = "Trezentos "; break;
                    case '4': centena = "Quatrocentos "; break;
                    case '5': centena = "Quinhentos "; break;
                    case '6': centena = "Seiscentos "; break;
                    case '7': centena = "Setecentos "; break;
                    case '8': centena = "Oitocentos "; break;
                    case '9': centena = "Novecentos "; break;

                }
                if(numero[1] == '1' && numero[2] == '0' && numero[3] == '0')
                {
                    centena = "Cem";
                }             
            }


            //ESCREVER POR EXTENSO DEZENA
            if (numero[2] == '1')
            {
                switch (numero.Substring(2))
                {
                    case "10": dezena = "e Dez"; break;
                    case "11": dezena = "e Onze"; break;
                    case "12": dezena = "e Doze"; break;
                    case "13": dezena = "e Treze"; break;
                    case "14": dezena = "e Quatorze"; break;
                    case "15": dezena = "e Quinze"; break;
                    case "16": dezena = "e Dezesseis"; break;
                    case "17": dezena = "e Dezessete"; break;
                    case "18": dezena = "e Dezoito"; break;
                    case "19": dezena = "e Dezenove"; break;
                }
            }
            else if (numero[2] != '0')
            {
                switch (numero[2])
                {
                    case '2': dezena = "e Vinte "; break;
                    case '3': dezena = "e Trinta "; break;
                    case '4': dezena = "e Quarenta "; break;
                    case '5': dezena = "e Cinquenta "; break;
                    case '6': dezena = "e Secenta "; break;
                    case '7': dezena = "e Setenta "; break;
                    case '8': dezena = "e Oitenta "; break;
                    case '9': dezena = "e Noventa "; break;
                }
            }


            //ESCREVER POR EXTENSO UNIDADE
            if (numero[3] != '0' && numero[2] != '1')
            {
                switch (numero[3])
                {
                    case '1': unidade = "e Um"; break;
                    case '2': unidade = "e Doiz"; break;
                    case '3': unidade = "e Três"; break;
                    case '4': unidade = "e Quatro"; break;
                    case '5': unidade = "e Cinco"; break;
                    case '6': unidade = "e Seis"; break;
                    case '7': unidade = "e Sete"; break;
                    case '8': unidade = "e Oito"; break;
                    case '9': unidade = "e Nove"; break;
                }
            }

            
            // CRIAR EXTENSO
            extenso = milhar + centena + dezena + unidade;

            // COMANDO DE SAIDA           
            if (extenso[0] == 'e')
            {
                extenso = extenso.Substring(2);
            }

            return extenso;           
        }


        public static string ValorParaExtenso2(decimal valor)
        {
            var wvalor = valor.ToString();
       
            string[] wunidade ={""," e um"," e dois"," e três"," e quatro"," e cinco"," e seis"," e sete"," e oito"," e nove"};
            string[] wdezes = {""," e onze"," e doze"," e treze"," e quatorze"," e quinze"," e dezesseis"," e dezessete"," e dezoito"," e dezenove"};
            string[] wdezenas = {""," e dez"," e vinte"," e trinta"," e quarenta"," e cinquenta"," e sessenta"," e setenta"," e oitenta"," e noventa"};
            string[] wcentenas = {""," e cento"," e duzentos"," e trezentos"," e quatrocentos"," e quinhentos"," e seiscentos"," e setecentos"," e oitocentos"," e novecentos"};
            string[] wplural = {" bilhões"," milhões"," mil",""};
            string[] wsingular = {" bilhão"," milhão"," mil",""};
            string wextenso = "";
            string wfracao;
            wvalor = wvalor.Replace("R$","");
            string wnumero = wvalor.Replace(",","").Trim();
            wnumero = wnumero.Replace(".","").PadLeft(14,'0');
            if(Int64.Parse(wnumero.Substring(0,12)) > 0){
                for(int i=0;i<4;i++){
                    wfracao = wnumero.Substring(i*3,3);
                    if(int.Parse(wfracao) != 0){
                        if(int.Parse(wfracao.Substring(0,3)) == 100) wextenso += " e cem";
                        else{
                            wextenso += wcentenas[int.Parse(wfracao.Substring(0,1))];
                            if(int.Parse(wfracao.Substring(1,2)) > 10 &&int.Parse(wfracao.Substring(1,2)) < 20)wextenso += wdezes[int.Parse(wfracao.Substring(2,1))];
                            else{
                                wextenso += wdezenas[int.Parse(wfracao.Substring(1,1))];
                                wextenso += wunidade[int.Parse(wfracao.Substring(2,1))];
                            }
                        }
                        if(int.Parse(wfracao) > 1) wextenso += wplural[i];
                        else wextenso += wsingular[i];
                    }
                }
                if(Int64.Parse(wnumero.Substring(0,12)) > 1) wextenso += " reais";
                else wextenso += " real";
            }
            wfracao = wnumero.Substring(12,2);
            if(int.Parse(wfracao) > 0){
                if(int.Parse(wfracao.Substring(0,2)) > 10 &&int.Parse(wfracao.Substring(0,2)) < 20) wextenso = wextenso  + wdezes[int.Parse(wfracao.Substring(1,1))];
                else{
                    wextenso += wdezenas[int.Parse(wfracao.Substring(0,1))];
                    wextenso += wunidade[int.Parse(wfracao.Substring(1,1))];
                }
                if(int.Parse(wfracao) > 1) wextenso += " centavos";
                else wextenso += " centavo";
            }
            if(wextenso != "") wextenso = wextenso.Substring(3,1).ToUpper()  + wextenso.Substring(4);            
            else wextenso = "Nada";

            return wextenso;
        

        }
    }
}