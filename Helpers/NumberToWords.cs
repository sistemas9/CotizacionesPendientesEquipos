using System;

namespace CotizacionesPendientesEquipos
{
  class NumberToWords
  {

    public static String ConvertToWords(String numb,String moneda)
    {
      String val = "", wholeNo = numb, points = "", andStr = "", pointStr = "";
      String endStr = "00/100 M.N.";
      String curr = String.Empty;
      String cents = String.Empty;
      if (moneda == "MXN")
      {
        curr = "pesos";
        cents = "M.N.";
      }
      try
      {
        int decimalPlace = numb.IndexOf(".");
        if (decimalPlace > 0)
        {
          wholeNo = numb.Substring(0, decimalPlace);
          points = numb.Substring(decimalPlace + 1);
          if (Convert.ToInt32(points) >= 0)
          {
            andStr = "";// "and";// just to separate whole numbers from points/cents  
            endStr = "";// "Paisa " + endStr;//Cents  
            //pointStr = ConvertDecimals(points);
            if (points.Length == 1)
            {
              points += "0";
            }
            pointStr = points+"/100 " + cents;
          }
        }
        val = String.Format("{0} {1}{2} {3}", ConvertWholeNumber(wholeNo).Trim()+" "+curr, andStr, pointStr, endStr);
      }
      catch { }
      return val;
    }
    private static String ConvertWholeNumber(String Number)
    {
      string word = "";
      try
      {
        bool beginsZero = false;//tests for 0XX
        bool isDone = false;//test if already translated
        double dblAmt = (Convert.ToDouble(Number));
        //if ((dblAmt > 0) && number.StartsWith("0"))
        if (dblAmt > 0)
        {//test for zero or digit zero in a nuemric
          beginsZero = Number.StartsWith("0");

          int numDigits = Number.Length;
          int pos = 0;//store digit grouping
          String place = "";//digit grouping name:hundres,thousand,etc...
          switch (numDigits)
          {
            case 1://ones' range

              word = ones(Number);
              isDone = true;
              break;
            case 2://tens' range
              word = tens(Number);
              if ((Convert.ToInt32(Number) % 30) == 0)
              {
                word = word.Replace("Treinta y", "Treinta");
              }
              if ((Convert.ToInt32(Number) % 40) == 0)
              {
                word = word.Replace("Cuarenta y", "Cuarenta");
              }
              if ((Convert.ToInt32(Number) % 50) == 0)
              {
                word = word.Replace("Cincuenta y", "Cincuenta");
              }
              if ((Convert.ToInt32(Number) % 60) == 0)
              {
                word = word.Replace("Sesenta y", "Sesenta");
              }
              if ((Convert.ToInt32(Number) % 70) == 0)
              {
                word = word.Replace("Setenta y", "Setenta");
              }
              if ((Convert.ToInt32(Number) % 80) == 0)
              {
                word = word.Replace("Ochenta y", "Ochenta");
              }
              if ((Convert.ToInt32(Number) % 90) == 0)
              {
                word = word.Replace("Noventa y", "Noventa");
              }
              isDone = true;
              break;
            case 3://hundreds' range
              pos = (numDigits % 3) + 1;
              place = " Ciento ";
              if ((Convert.ToInt32(Number) % 100) == 0)
              {
                place = " Cien ";
              }              
              break;
            case 4://thousands' range
            case 5:
            case 6:
              pos = (numDigits % 4) + 1;
              place = " Mil ";
              break;
            case 7://millions' range
            case 8:
            case 9:
              pos = (numDigits % 7) + 1;
              place = " Un Millon ";
              break;
            case 10://Billions's range
            case 11:
            case 12:

              pos = (numDigits % 12) + 1;
              place = " Billiones ";
              break;
            //add extra case options for anything above Billion...
            default:
              isDone = true;
              break;
          }
          if (!isDone)
          {//if transalation is not done, continue...(Recursion comes in now!!)
            if (Number.Substring(0, pos) != "0" && Number.Substring(pos) != "0")
            {
              try
              {
                word = ConvertWholeNumber(Number.Substring(0, pos)) + place + ConvertWholeNumber(Number.Substring(pos));
              }
              catch { }
            }
            else
            {
              word = ConvertWholeNumber(Number.Substring(0, pos)) + ConvertWholeNumber(Number.Substring(pos));
            }

            //check for trailing zeros
            //if (beginsZero) word = " and " + word.Trim();
          }
          //ignore digit grouping names
          if (word.Trim().Equals(place.Trim())) word = "";
        }
      }
      catch { }
      ////////filtros especiales//////////
      ///
      word = word.Replace("Veinte Uno", "Veintiun");
      word = word.Replace("Veinte Dos", "Veintidos");
      word = word.Replace("Veinte Tres", "Veintitres");
      word = word.Replace("Veinte Cuatro", "Veinticuatro");
      word = word.Replace("Veinte Cinco", "Veinticinco");
      word = word.Replace("Veinte Seis", "Veintiseis");
      word = word.Replace("Veinte Siete", "Veintisiete");
      word = word.Replace("Veinte Ocho", "Veintiocho");
      word = word.Replace("Veinte Nueve", "Veintinueve");

      word = word.Replace("Uno Ciento", "Ciento");
      word = word.Replace("Uno Cien", "Cien");
      word = word.Replace("Dos Ciento", "Doscientos");
      word = word.Replace("Dos Cien", "Doscientos");
      word = word.Replace("Tres Ciento", "Trescientos");
      word = word.Replace("Tres Cien", "Trescientos");
      word = word.Replace("Cuatro Ciento", "Cuatrocientos");
      word = word.Replace("Cuatro Cien", "Cuatrocientos");
      word = word.Replace("Cinco Ciento", "Quinientos");
      word = word.Replace("Cinco Cien", "Quinientos");
      word = word.Replace("Seis Ciento", "Seiscientos");
      word = word.Replace("Seis Cien", "Seiscientos");
      word = word.Replace("Siete Ciento", "Setecientos");
      word = word.Replace("Siete Cien", "Setecientos");
      word = word.Replace("Ocho Ciento", "Ochocientos");
      word = word.Replace("Ocho Cien", "Ochocientos");
      word = word.Replace("Nueve Ciento", "Novecientos");
      word = word.Replace("Nueve Cien", "Novecientos");

      word = word.Replace("Treinta y Uno Mil", "Treintaiunmil");
      word = word.Replace("Cuarenta y Uno Mil", "Cuarentaiunmil");
      word = word.Replace("Ciancuenta y Uno Mil", "Ciancuentaiunmil");
      word = word.Replace("Sesenta y Uno Mil", "Sesentaiunmil");
      word = word.Replace("Setenta y Uno Mil", "Setentaiunmil");
      word = word.Replace("Ochenta y Uno Mil", "Ochentaiunmil");
      word = word.Replace("Noventa y Uno Mil", "Noventaiunmil");

      word = word.Replace("Trescientos Uno Mil", "Trescientosunmil");
      word = word.Replace("Cuatrocientos Uno Mil", "Cuatrocientosunmil");
      word = word.Replace("Quinientos Uno Mil", "Quinientosunmil");
      word = word.Replace("Seiscientos Uno Mil", "Seiscientosunmil");
      word = word.Replace("Setecientos Uno Mil", "Setecientosunmil");
      word = word.Replace("Ochocientos Uno Mil", "Ochocientosunmil");
      word = word.Replace("Novecientos Uno Mil", "Novencientosunmil");

      word = word.Replace("Uno Un Millon", "Un Millon");
      word = word.Replace("Dos Un Millon", "Dos Millones");
      word = word.Replace("Tres Un Millon", "Tes Millones");
      word = word.Replace("Cuatro Un Millon", "Cuatro Millones");
      word = word.Replace("Cinco Un Millon", "Cinco Millones");
      word = word.Replace("Seis Un Millon", "Seis Millones");
      word = word.Replace("Siete Un Millon", "Siete Millones");
      word = word.Replace("Ocho Un Millon", "Ocho Millones");
      word = word.Replace("Nueve Un Millon", "Nueve Millones");

      word = word.Replace("Uno Mil", "Mil");
      return word.Trim();
    }

    private static String tens(String Number)
    {
      int _Number = Convert.ToInt32(Number);
      String name = null;
      switch (_Number)
      {
        case 10:
          name = "Diez";
          break;
        case 11:
          name = "Once";
          break;
        case 12:
          name = "Doce";
          break;
        case 13:
          name = "Trece";
          break;
        case 14:
          name = "Catorce";
          break;
        case 15:
          name = "Quince";
          break;
        case 16:
          name = "Dieciseis";
          break;
        case 17:
          name = "Diecisiete";
          break;
        case 18:
          name = "Dieciocho";
          break;
        case 19:
          name = "Diecinueve";
          break;
        case 20:
          name = "Veinte";
          break;
        case 30:
          name = "Treinta y";
          break;
        case 40:
          name = "Cuarenta y";
          break;
        case 50:
          name = "Cincuenta y";
          break;
        case 60:
          name = "Sesenta y";
          break;
        case 70:
          name = "Setenta y";
          break;
        case 80:
          name = "Ochenta y";
          break;
        case 90:
          name = "Noventa y";
          break;
        default:
          if (_Number > 0)
          {
            name = tens(Number.Substring(0, 1) + "0") + " " + ones(Number.Substring(1));
          }
          break;
      }
      return name;
    }
    private static String ones(String Number)
    {
      int _Number = Convert.ToInt32(Number);
      String name = "";
      switch (_Number)
      {

        case 1:
          name = "Uno";
          break;
        case 2:
          name = "Dos";
          break;
        case 3:
          name = "Tres";
          break;
        case 4:
          name = "Cuatro";
          break;
        case 5:
          name = "Cinco";
          break;
        case 6:
          name = "Seis";
          break;
        case 7:
          name = "Siete";
          break;
        case 8:
          name = "Ocho";
          break;
        case 9:
          name = "Nueve";
          break;
      }
      return name;
    }
    private static String ConvertDecimals(String number)
    {
      String cd = "", digit = "", engOne = "";
      for (int i = 0; i < number.Length; i++)
      {
        digit = number[i].ToString();
        if (digit.Equals("0"))
        {
          engOne = "Zero";
        }
        else
        {
          engOne = ones(digit);
        }
        cd += " " + engOne;
      }
      return cd;
    }
  }
}
