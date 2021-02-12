Attribute VB_Name = "Background"
Option Explicit

'Here we formulate how the Algorithm will work in the program

Public Illness, Symptoms As Currency

Sub Symptoms_Selection()
  If chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked _
                       And chckIndigestion.Value = vbChecked And chckSputum.Value = vbChecked Then
        Symptoms = Fever + Cough + Anorexia + Myalgia + Dyspenea + Sputum
     
           ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked _
                       And chckIndigestion.Value = vbChecked Then
                  Symptoms = Fever + Cough + Anorexia + Myalgia + Dyspenea
     
 
        ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked Then
           Symptoms = Fever + Cough + Anorexia + Myalgia
 
              ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckWLoss.Value = vbChecked Then
                  Symptoms = Fever + Cough + Anorexia
     
 
                     ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked _
                       And chckSputum.Value = vbChecked Then
                        Symptoms = Fever + Cough + Anorexia + Myalgia + Sputum
     
 
                         ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckSputum.Value = vbChecked Then
                            Symptoms = Fever + Cough + Anorexia + Sputum
     
  
                               ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckIndigestion.Value = vbChecked Then
                                  Symptoms = Fever + Cough + Anorexia + Dyspenea
     

 
                                        ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckSputum.Value = vbChecked Then
                                           Symptoms = Fever + Cough + Sputum
     
  
                                                ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckIndigestion.Value = vbChecked Then
                                                   Symptoms = Fever + Cough + Dyspenea
     
 
 
                                                    ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked Then
                                                        Symptoms = Fever + Cough + Myalgia
 
                                                           ElseIf chckFever.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                              Symptoms = Fever + Sputum
     
 
                                                                 ElseIf chckFever.Value = vbChecked And chckIndigestion.Value = vbChecked Then
                                                                     Symptoms = Fever + Dyspenea
     
  
                                                                       ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked _
                                                                            And chckIndigestion.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                                             Symptoms = Fever + Cough + Anorexia + Myalgia + Dyspenea + Sputum
     

                         
                                                                                  ElseIf chckFever.Value = vbChecked And chckMuscle.Value = vbChecked Then
                                                                                    Symptoms = Fever + Myalgia
                                                                                    
                                                                                    
                                                                                         ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked _
                                                                                            And chckIndigestion.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                                                            Symptoms = Fever + Cough + Anorexia + Myalgia + Dyspenea + Sputum
     
 
                                                                                                 ElseIf chckFever.Value = vbChecked And chckWLoss.Value = vbChecked Then
                                                                                                      Symptoms = Fever + Anorexia
    
 
                                                                                                      ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked _
                                                                                                          And chckIndigestion.Value = vbChecked And chckSputum.Value = vbChecked Then
  
                                                                                                               Symptoms = Fever + Cough + Anorexia + Myalgia + Dyspenea + Sputum
     

 
                                                                                                                    ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked Then
                                                                                                                        Symptoms = Fever + Cough
     

 
                                                                                                                         ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked _
                                                                                                                             And chckIndigestion.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                                                                                               Symptoms = Fever + Cough + Anorexia + Myalgia + Dyspenea + Sputum
     
 
                                                                                                                                   ElseIf chckFever.Value = vbChecked Then
                                                                                                                                     Symptoms = Fever
     
                                                                                                                                        ElseIf chckFever.Value = vbChecked And chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked _
                                                                                                                                             And chckIndigestion.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                                                                                                                Symptoms = Fever + Cough + Anorexia + Myalgia + Dyspenea + Sputum
     
 
                                                                                                                                                  ElseIf chckCough.Value = vbChecked Then
                                                                                                                                                  Symptoms = Cough
                                                                                                                                                  
 
                                                                                                                                                       ElseIf chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked _
                                                                                                                                                       And chckIndigestion.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                                                                                                                          Symptoms = Cough + Anorexia + Myalgia + Dyspenea + Sputum
     
                                                                                                                                                                ElseIf chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked _
                                                                                                                                                                  And chckIndigestion.Value = vbChecked Then
                                                                                                                                                                    Symptoms = Cough + Anorexia + Myalgia + Dyspenea
     
                                                                                                                                                                    ElseIf chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked _
                                                                                                                                                                        And chckIndigestion.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                                                                                                                                           Symptoms = Cough + Anorexia + Myalgia + Sputum
                                                                                                                                                                        
                                                                                                                                                                              ElseIf chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                                                                                                                                                 Symptoms = Cough + Anorexia + Sputum
                                                                                                                                                                                 
                                                                                                                                                                                   ElseIf chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckIndigestion.Value = vbChecked Then
                                                                                                                                                                                      Symptoms = Cough + Anorexia + Dyspenea
                                                                                                                                                                                      
                                                                                                                                                                                      ElseIf chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked Then
                                                                                                                                                                                          Symptoms = Cough + Anorexia + Myalgia
     
 
                                                                                                                                                                                             ElseIf chckCough.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                                                                                                                                                                 Symptoms = Cough + Sputum
 
                                                                                                                                                                                                     ElseIf chckCough.Value = vbChecked And chckIndigestion.Value = vbChecked Then
                                                                                                                                                                                                        Symptoms = Cough + Dyspenea
     
                                                                                                                                                                                                           ElseIf chckCough.Value = vbChecked And chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked And chckIndigestion.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                                                                                                                                                                               Symptoms = Cough + Anorexia + Myalgia + Dyspenea + Sputum
 
                                                                                                                                                                                                                    ElseIf chckCough.Value = vbChecked And chckMuscle.Value = vbChecked Then
                                                                                                                                                                                                                       Symptoms = Cough + Myalgia
                                                                                                                                                                                                                       
 
                                                                                                                                                                                                                             ElseIf chckCough.Value = vbChecked And chckWLoss.Value = vbChecked Then
                                                                                                                                                                                                                                 Symptoms = Cough + Anorexia
     
 
 
                                                                                                                                                                                                                                   ElseIf chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked And chckIndigestion.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                                                                                                                                                                                                        Symptoms = Anorexia + Myalgia + Dyspenea + Sputum
     
     
                                                                                                                                                                                                                                             ElseIf chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                                                                                                                                                                                                                  Symptoms = Anorexia + Myalgia + Sputum
     
                                                                                                                                                                                                                                                       ElseIf chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked And chckIndigestion.Value = vbChecked Then
                                                                                                                                                                                                                                                           Symptoms = Anorexia + Myalgia + Dyspenea
                                                                                                                                                                                                                                                           
                                                                                                                                                                                                                                                           ElseIf chckWLoss.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                                                                                                                                                                                                                              Symptoms = Anorexia + Sputum
                                                                                                                                                                                                                                                               
  ElseIf chckWLoss.Value = vbChecked And chckIndigestion.Value = vbChecked Then
   Symptoms = Anorexia + Dyspenea
     
 
    ElseIf chckWLoss.Value = vbChecked And chckMuscle.Value = vbChecked Then
       Symptoms = Anorexia + Myalgia
     
 
          ElseIf chckWLoss.Value = vbChecked Then
              Symptoms = Anorexia
     
                 ElseIf chckMuscle.Value = vbChecked And chckIndigestion.Value = vbChecked And chckSputum.Value = vbChecked Then
                      Symptoms = Myalgia + Dyspenea + Sputum
     
 
                            ElseIf chckMuscle.Value = vbChecked And chckSputum.Value = vbChecked Then
                                  Symptoms = Myalgia + Sputum
     
 
                                           ElseIf chckMuscle.Value = vbChecked And chckIndigestion.Value = vbChecked Then
                                                Symptoms = Myalgia + Dyspenea
     
 
                                                        ElseIf chckMuscle.Value = vbChecked Then
                                                           Symptoms = Myalgia
     
 
                                                              ElseIf chckIndigestion.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                                   Symptoms = Dyspenea + Sputum
     
 
                                                                       ElseIf chckMuscle.Value = vbChecked And chckIndigestion.Value = vbChecked And chckSputum.Value = vbChecked Then
                                                                           Symptoms = Dyspenea
     
 
                                                                           ElseIf chckSputum.Value = vbChecked Then
                                                                                Symptoms = Sputum
                                                                                
    Else:
         Symptoms = 0
         
  End If
      
End Sub




Sub Illness_Selection()
 If chckHIV = vbChecked And chckCVD = vbChecked And chckDiabetes = vbChecked Then
    Illness = HIV + CVD + Diabetes
    
    ElseIf chckCVD = vbChecked And chckDiabetes = vbChecked Then
      Illness = CVD + Diabetes
    
       ElseIf chckHIV = vbChecked And chckCVD = vbChecked Then
         Illness = HIV + CVD

         ElseIf chckHIV = vbChecked And chckDiabetes = vbChecked Then
           Illness = HIV + Diabetes
             
             ElseIf chckHIV = vbChecked Then
               Illness = HIV
               
                ElseIf chckDiabetes = vbChecked Then
                   Illness = Diabetes
                   
                     ElseIf chckCVD = vbChecked Then
                       Illness = CVD
                       
    Else:
        
        Illness = 0
        
 End If
 
          
End Sub


