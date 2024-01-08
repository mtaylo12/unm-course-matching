import pandas as pd
from mip import *
import argparse

parser = argparse.ArgumentParser(description='Process input file and write to output file')
parser.add_argument('input_file', help='Input file name')
parser.add_argument('output_file', help='Output file name')

args = parser.parse_args()
input = args.input_file
output = args.output_file



xls_file = input
file = pd.ExcelFile(xls_file)

sheet_sections =  pd.read_excel(file, sheet_name ="sections")     #row for each section: course number, day, time
sheet_coordpref =  pd.read_excel(file, sheet_name ="coordpref")   #row for each instructor: instructor name, course1, ...
sheet_instpref =  pd.read_excel(file, sheet_name ="instpref")     #row for each instructor: instructor name, course1, ...
sheet_instructors = pd.read_excel(file, sheet_name = "instructors")  #row for each instructor: instructor name, max, min, unavail (T), unavail (M)


def solve_model(instructors_df, sections_df, coord_preferences_df, inst_preferences_df, output_file):
    section_codes = sections_df.index.values.tolist()
    instructor_codes = instructors_df.index.values.tolist()
  
    sum_pref = pd.concat([coord_preferences_df, inst_preferences_df]).groupby(['Instructor']).sum().reset_index()

    ########### INPUT FORMAT TESTING ###########
    
    
    ########### MODEL INITIALIZATION ###########
    m = Model(sense=MAXIMIZE, solver_name=CBC)
    
    # decision variables: x[i][j] denotes if lecturer[i] teaches section[j]
    x = [[m.add_var(var_type=BINARY) for j in section_codes] for i in instructor_codes]

    # decision variables: y[i][d] denotes if lecturer i teaches on day d where d is in ['M', 'T']
    day_codes = [0, 1]
    days = ['M', 'T']
    y = [[m.add_var(var_type=BINARY) for j in day_codes] for i in instructor_codes]

    # decision variables: z[i][t] denotes if lecturer i teaches in morning, midday, or afternoon
    # time_codes = [0, 1, 2]
    # z = [[m.add_var(var_type=BINARY) for j in time_codes] for i in instructor_codes]

    ########### BASIC CONSTRAINTS ###########
    # constraint: each section has exactly 1 lecturer
    for j in section_codes:
        m += xsum(x[i][j] for i in instructor_codes) == 1
    
    # constraint: each lecturer teaches at least their min
    for i in instructor_codes:
        m += xsum(5/3 * x[i][j]  if sections_df['Course'].values[j] == 1250 else x[i][j] for j in section_codes) >= instructors_df['Min'].values[i]

    # constraint: each lecturer teaches at most their max
    for i in instructor_codes:
        m += xsum(5/3 * x[i][j] if sections_df['Course'].values[j] == 1250 else x[i][j] for j in section_codes) <= instructors_df['Max'].values[i]

    # # setup y based on x values to denote if instructor i teaches on day d
    for i in instructor_codes:
        for day_code in day_codes:
            for j in section_codes:
                if sections_df['Day'].values[j] == days[day_code] or sections_df['Day'].values[j] == 'MT':
                    m += y[i][day_code] >= x[i][j]

    # # setup z based on x values to denote if i teaches in morning or afternoon
    # for i in instructor_codes:
    #     for j in section_codes:
    #         if sections_df['Time'].values[j] < 12:
    #             m += z[i][0] >= x[i][j]
            
    #         if sections_df['Time'].values[j] < 2 and sections_df['Time'].values[j] > 9:
    #             m += z[i][1] >= x[i][j]
            
    #         if sections_df['Time'].values[j] > 1:
    #             m += z[i][2] >= x[i][j]


    # ########### PARTICULAR CONSTRAINTS ###########
    # constraint: for calculus (1512, 1522, 2531) an instructor can only teach one of each section
    for i in instructor_codes:
        for course_number in [1512, 1522, 2531]:
            m += xsum(x[i][j] for j in section_codes if sections_df['Course'].values[j] == course_number) <= 1

    # constraint: an instructor can teach a max of two sections of a given course
    for i in instructor_codes:
        for course_number in sections_df['Course'].unique():
            m += xsum(x[i][j] for j in section_codes if sections_df['Course'].values[j] == course_number) <= 2

    # # constraint: some instructors in at least one calculus class
    for inst in ['TimBerkopec', 'PatrickDenne', 'KhalidIfzarene', 'KarenChampine']:
        i = instructors_df.index[instructors_df['Instructor'] == inst].tolist()[0]
        print(sections_df['Course'].values[10] in [1512, 1522, 2531])
        m += xsum(x[i][j] for j in section_codes if (sections_df['Course'].values[j] in [1512, 1522, 2531])) >= 1

    # constraint: some instructors in at least one 1250 course
    for inst in ['HuynhDinh', 'PatrickDenne']:
        i = instructors_df.index[instructors_df['Instructor'] == inst].tolist()[0]
        m += sum(x[i][j] for j in section_codes if sections_df['Course'].values[j]==1250) >= 1

    # TODO: turn into a preference not a constraint
    for inst in ['TimBerkopec']:
        i = instructors_df.index[instructors_df['Instructor'] == inst].tolist()[0]
        for course_number in sections_df['Course'].unique():
            m += xsum(x[i][j] for j in section_codes if sections_df['Course'].values[j] == course_number) <= 1
        
    # ########### COMPLEX CONSTRAINTS ###########
    # constraint: each lecturer can only teach courses matching ability
    for i in instructor_codes:
        m += xsum((x[i][j] for j in section_codes if coord_preferences_df.iloc[i][sections_df['Course'].values[j]]==0)) == 0

    # # constraint: a lecturer can't teach two courses at the same time
    for i in instructor_codes:
        for t in range(0, 24):
            for d in ["M", "T"]:
                common_time_section_codes = filter(lambda sec : ((sections_df["Day"].values[sec] == d or sections_df["Day"].values[sec] == 'MT') and sections_df["Time"].values[sec] == t), section_codes)
                m += xsum(x[i][j] for j in common_time_section_codes) <= 1

    # # constraint: a lecturer can only teach section_codes for which they are available (time constraints)
    for i in instructor_codes:
        m += xsum(x[i][j] for j in section_codes if 
                ((int(sections_df['Time'].values[j]) in eval(instructors_df['unavailable (T)'].values[i])) and ((sections_df['Day'].values[j] == 'T') or sections_df['Day'].values[j] == 'MT'))
                or 
                ((int(sections_df['Time'].values[j]) in eval(instructors_df['unavailable (M)'].values[i])) and ((sections_df['Day'].values[j] == 'M') or sections_df['Day'].values[j] == 'MT'))) == 0

    ########### OBJECTIVE FUNCTION ###########
    m.objective = maximize(xsum(x[i][j] * inst_preferences_df.iloc[i][sections_df['Course'].values[j]]for i in instructor_codes for j in section_codes)
                            + xsum(x[i][j] * 2 * coord_preferences_df.iloc[i][sections_df['Course'].values[j]]for i in instructor_codes for j in section_codes)
                            - xsum(y[i][d] for i in instructor_codes for d in day_codes))
                            #- xsum(z[i][d] for i in instructor_codes for d in time_codes)) 


    ########### SOLVE ###########
    m.optimize()

    ########### PROCESS RESULTS ###########
    print("########### RESULTS ###########")
    print("Objective value: ", m.objective_value)
    print("Number of solutions: ", m.num_solutions)
    print("Number of sections: ", len(section_codes))
    print("Max possible: ", instructors_df['Max'].sum())
    print("Min possible: ", instructors_df['Min'].sum())
    if m.num_solutions:
        print("Writing solution to", output)
        resdf = pd.DataFrame(columns=['Instructor', 'Course', 'Day', 'Time'])

        solution = np.zeros([len(instructor_codes), len(section_codes)])
        for i in instructor_codes:
            for j in section_codes:
                if x[i][j].x == 1:
                    resdf.loc[j] = [instructors_df["Instructor"].values[i], sections_df["Course"].values[j], sections_df["Day"].values[j], sections_df["Time"].values[j]]
            

        resdf.to_excel(output_file)


     
    else:
        print("No solution found.")




#print(type(eval(sheet_instructors['unavailable (M)'].values[1])[0]))
#print(type(int(sheet_sections['Time'].values[0])))
solve_model(sheet_instructors, sheet_sections, sheet_coordpref, sheet_instpref, output)