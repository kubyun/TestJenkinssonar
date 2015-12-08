package bamboo_test.bamboo_test;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        System.out.println( "Hello World!" );
        System.out.println( "Ah!!!!! junit!!!!!" );

        
        //testtesttest
        
        for(int i=1; i<=9; i++) {
            for(int j=2; j<=5; j++) {
                System.out.printf("%d * %d = %2d", j, i, (j*i));
                System.out.print("\t");
            }
            System.out.println(); 
        }
 
        System.out.println();
 
        for(int i=1; i<=8; i++) {
            for(int j=6; j<=8; j++) {
                System.out.printf("%d * %d = %2d", j, i, (j*i));
                System.out.print("\t");
            }
            System.out.println();
            System.out.print("test maven ver 0.1.1");
           
        }      
        
        
        for(int i=1; i<=8; i++) {
            for(int j=6; j<=8; j++) {
                System.out.printf("%d * %d = %2d", j, i, (j*i));
                System.out.print("\t");
            }
            System.out.println();
            System.out.print("test maven ver 0.1.2");
           
        }      
       
       
        
        
    }
}
