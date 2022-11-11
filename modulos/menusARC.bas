Attribute VB_Name = "menusARC"

Sub pl_deshabilita_menu()
    MDI_Administrador_contable.mnu_archivo.Enabled = True
    MDI_Administrador_contable.Mnu_consulta.Enabled = False
    MDI_Administrador_contable.mnuResponsabilidades.Enabled = False
    MDI_Administrador_contable.mnuProcesos.Enabled = False
End Sub

Sub pl_habilita_menu()
    MDI_Administrador_contable.mnu_archivo.Enabled = True
    MDI_Administrador_contable.Mnu_consulta.Enabled = True
    MDI_Administrador_contable.mnuResponsabilidades.Enabled = True
    MDI_Administrador_contable.mnuProcesos.Enabled = True
End Sub
