import java.util.ArrayList;
import java.util.List;

/**
 * @author wangjs6
 * @version 1.0
 * @Description:
 * @date: 2019/8/5 9:01
 */
public class Test {

    @org.junit.Test
    public void test(){
        int task = 5;
        int count = 20;
        if(task > count){
            System.out.println(getList(0,count-1));
            return;
        }
        int everyTask = count / task;
        int last = count % task;

        for(int i =0;i<task;i++){
            System.out.print(i+":");
            if(i == task - 1){
                if(last != 0){
                    System.out.println(getList(i*everyTask,count-1));
                }else{
                    System.out.println(getList(i*everyTask,i*everyTask+everyTask-1));
                }
            }else{
                System.out.println(getList(i*everyTask,i*everyTask+everyTask-1));
            }
        }
    }
    private List<Integer> getList(int start,int end){
        List<Integer> list = new ArrayList<Integer>();
        for(int i = start;i <= end;i++){
            list.add(i);
        }
        return list;
    }

}
