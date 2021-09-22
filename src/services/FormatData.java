package services;

import bean.ClassRating;
import bean.RoutineInspection;

import java.text.DecimalFormat;
import java.util.*;

public class FormatData {

    public RoutineInspection SummaryInspection(ArrayList<RoutineInspection> routineInspections) {

        if (routineInspections == null) {
            return null;
        }

        ArrayList<ClassRating> classRatings = new ArrayList<>();
        SortedMap<String, double[]> summary = new TreeMap<>();
        for (RoutineInspection routineInspection : routineInspections) {
            for (ClassRating classRating : routineInspection.cr) {
                double[] rating = summary.get(classRating.className);
                if (rating == null) {
                    rating = new double[6];
                }
                rating[0] += classRating.moral;
                rating[1] += classRating.read;
                rating[2] += classRating.wisdom;
                rating[3] += classRating.health;
                rating[4] += classRating.art;
                rating[5] += classRating.practice;
                summary.put(classRating.className, rating);
            }
        }

        for (Map.Entry<String, double[]> entry : summary.entrySet()) {
            ClassRating classRating = new ClassRating();
            classRating.className = entry.getKey();
            classRating.moral = entry.getValue()[0];
            classRating.read = entry.getValue()[1];
            classRating.wisdom = entry.getValue()[2];
            classRating.health = entry.getValue()[3];
            classRating.art = entry.getValue()[4];
            classRating.practice = entry.getValue()[5];

            if (classRating.moral >= 0) {
                classRating.star++;
            }
            if (classRating.read >= 0) {
                classRating.star++;
            }
            if (classRating.wisdom >= 0) {
                classRating.star++;
            }
            if (classRating.health >= 0) {
                classRating.star++;
            }
            if (classRating.art >= 0) {
                classRating.star++;
            }
            if (classRating.practice >= 0) {
                classRating.star++;
            }

            classRatings.add(classRating);
        }

        return new RoutineInspection(classRatings, "汇总");
    }

}
