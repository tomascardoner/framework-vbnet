
' Written by Kevin Ng.
' This full article can be found @ http://kevinhng86.iblog.website

Friend Class CS_Conversion

 public static void main(String[] args)
 {
   // Test case
    String a = "11111000111100010";
    Integer b = bindecZ(a);
    String c = decbinZ(b);

    System.out.println(a);
    System.out.println(b);
    System.out.println(c);

 }


    Public Function BinaryToDecimal(ByVal invalue As String) As Integer
        '* Formula: Starting at the right of the string.
        '*          Treat the binaray digit 0 as if it is a decimal 0, and 1 as if it is a decimal 1.
        '*          Each binary digit is times to 2 by the power of their position. And the first one being 0.
        '*          Sum the total of each digit times 2 by the power of their position together and that is the decimal of that binary string.    
        '*
        '*    Example: Converting binary 1111 to decimal    
        '*    Step 1: (1x2e3) + (1x2e2) + (1x2e1) + (1x2e0)    
        '*    Step 2:    8    +    4    +    2         1
        '*       Answer:                  15

        Dim times As Integer = 1
        Dim output As Integer = 0

        Do While (invalue.Length Mod 8) <> 0
            invalue = "0" + invalue
        Loop

        For i As Integer = invalue.Length - 1 To 0 Step -1
            output += (((Integer.parseInt(String.valueOf(invalue.charAt(i)))) * times))
            times *= 2
        Next i

        Return output
    End Function

    public static String decbinZ(int invalue){
    /*
     *      Formula: Divide the decimal number by 2 until the decimal number become 1.
     *               Each time the decimal number is divide, a binary digit is produce base on two condition. 
     *               Write down the binary digit from each division from first being the right most.
     *               If the decimal number is divisble by 2 we store a 0 binary digit.
     *               If the decimal number is not divisble by 2, minus one of the decimal number, divide it by two and write down a 1 for the binary digit. 
     *               When the decimal number become 1, that is the final binary digit and that is a 1 binary digit.
     *    
     *    Example: Convert 254 decimal to binary
     *               
     *             Equation             Binary Digit     Binary Digit After Equation         
     *  Step 1 : 254 / 2       =  127       0                     0
     *  Step 2 : (127 - 1) / 2 =  63        1                   1 0
     *  Step 3 : (63 - 1 ) / 2 =  31        1                 1 1 0
     *  Step 4 : (31 - 1 ) / 2 =  15        1               1 1 1 0
     *  Step 5 : (15 - 1 ) / 2 =   7        1             1 1 1 1 0
     *  Step 6 : ( 7 - 1 ) / 2 =   3        1           1 1 1 1 1 0
     *  Step 7 : ( 3 - 1 ) / 2 =   1        1         1 1 1 1 1 1 0 
     *  Step 8 :    1  move to binary       1       1 1 1 1 1 1 1 0
     *         Answer : 254 decimal value is 11111110 in binary value.     
     */
        String output = "";
        if  (invalue == 0){
            return "0000";
        }
        while (invalue > 0){
            if (invalue == 1){
                invalue = invalue - 1;
                output  = "1" + output;
            } else if (invalue % 2 == 0){
                invalue = invalue / 2;
                output  = "0" + output;
            } else if (invalue % 2 != 0){
                invalue = (invalue - 1) / 2;
                output = "1" + output;
            }
        // This two line below is the proper way of writting all the if else code block into two line. Remove the if else if use this below two line.
        //    output =  (invalue == 1)? "1" + output : (invalue % 2 == 0)? "0" + output : "1" + output;
        //    invalue = (invalue % 2 == 0)? (invalue / 2): (invalue - 1) / 2;
        }


        // This code block is when the full set of 8 binary digit must be return.
        //while ( output.length() % 8 != 0){
        //    output = "0" + output;
        //}        
        return output;
    }

End Class