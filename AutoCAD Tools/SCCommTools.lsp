(defun C:CTImportUCSList ()
	(vl-load-com)
	(setq oApp (vlax-get-acad-object))
	(vla-RunMacro oApp "SCCommTools.AcToolsUCS.CTImportUCSList")
	(princ)
)

(defun C:CTLabelHighlightOrphaned ()
	(vl-load-com)
	(setq oApp (vlax-get-acad-object))
	(vla-RunMacro oApp "SCCommTools.acToolsLabels.CTLabelHighlightOrphaned")
	(princ)
)

(defun C:CTLightHole ()
	(vl-load-com)
	(setq oApp (vlax-get-acad-object))
	(vla-RunMacro oApp "SCCommTools.AcToolsLightHole.CTLightHole")
	(princ)
)

(defun C:CTRenameUCSList ()
	(vl-load-com)
	(setq oApp (vlax-get-acad-object))
	(vla-RunMacro oApp "SCCommTools.AcToolsUCS.CTRenameUCSList")
	(princ)
)

(defun C:CTSpaceBomb ()
	(vl-load-com)
	(setq oApp (vlax-get-acad-object))
	(vla-RunMacro oApp "SCCommTools.AcToolsSpaceBomb.CTSpaceBomb")
	(princ)
)


(defun C:CTMDGroupObjects ()
	(vl-load-com)
	(setq oApp (vlax-get-acad-object))
	(vla-RunMacro oApp "SCCommTools.AcToolsMDGroupObjs.CTMDGroupObjects")
	(princ)
)

(defun C:CTMDPostProcess ()
	(vl-load-com)
	(setq oApp (vlax-get-acad-object))
	(vla-RunMacro oApp "SCCommTools.AcToolsMDPostProcess.CTMDPostProcess")
	(princ)
)