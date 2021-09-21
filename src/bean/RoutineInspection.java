package bean;

import java.util.ArrayList;

public class RoutineInspection {

    public ArrayList<ClassRating> cr;
    public String fileName;

    public RoutineInspection() {
    }

    public RoutineInspection(ArrayList<ClassRating> classRating, String fileName) {
        this.cr = classRating;
        this.fileName = fileName;
    }


}
