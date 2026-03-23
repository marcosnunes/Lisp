(defun AGM:DegToRad (deg)
  (* pi (/ deg 180.0))
)

(defun AGM:NormalizeNumText (txt)
  (vl-string-translate "," "." (vl-string-trim " " txt))
)

(defun AGM:GetDecimalInput (msg / raw val)
  (setq val nil)
  (while (null val)
    (setq raw (getstring T msg))
    (if (or (null raw) (= (vl-string-trim " " raw) ""))
      (princ "\nValor invalido. Digite um numero em graus decimais.")
      (progn
        (setq raw (AGM:NormalizeNumText raw))
        (setq val (distof raw 2))
        (if (null val)
          (princ "\nFormato invalido. Exemplo: -23.55052")
        )
      )
    )
  )
  val
)

(defun AGM:ParseLatLonPair (raw / txt sep lat-txt lon-txt lat lon)
  (if raw
    (progn
      (setq txt (vl-string-trim " " raw))
      (setq sep (vl-string-search "," txt))
      (if sep
        (progn
          (setq lat-txt (vl-string-trim " " (substr txt 1 sep)))
          (setq lon-txt (vl-string-trim " " (substr txt (+ sep 2))))
          (setq lat (distof (AGM:NormalizeNumText lat-txt) 2))
          (setq lon (distof (AGM:NormalizeNumText lon-txt) 2))
          (if (and lat lon)
            (list lat lon)
            nil
          )
        )
        nil
      )
    )
    nil
  )
)

(defun AGM:LatLonToUTM (lat lon / semi-major f e2 ep2 k0 zone lon0 lat-abs phi lam lam0 sinp cosp tanp N tan2 C dlam-cos M e4 e6 m1 m2 m3 m4
                             easting northing hemisphere epsg utm-code)
  ;; SIRGAS2000 uses GRS80 ellipsoid parameters
  (setq semi-major 6378137.0)
  (setq f (/ 1.0 298.257222101))
  (setq e2 (* f (- 2.0 f)))
  (setq ep2 (/ e2 (- 1.0 e2)))
  (setq k0 0.9996)

  (setq zone (1+ (fix (/ (+ lon 180.0) 6.0))))
  (if (< zone 1) (setq zone 1))
  (if (> zone 60) (setq zone 60))

  (setq lon0 (- (* zone 6.0) 183.0))

  (setq hemisphere (if (< lat 0.0) "S" "N"))
  (setq lat-abs (abs lat))

  ;; Para estabilidade numerica, calcula a projeção com latitude positiva
  ;; e aplica false northing no fim para hemisferio sul.
  (setq phi (AGM:DegToRad lat-abs))
  (setq lam (AGM:DegToRad lon))
  (setq lam0 (AGM:DegToRad lon0))

  (setq sinp (sin phi))
  (setq cosp (cos phi))
  (setq tanp (tan phi))

  (setq N (/ semi-major (sqrt (- 1.0 (* e2 sinp sinp)))))
  (setq tan2 (* tanp tanp))
  (setq C (* ep2 cosp cosp))
  (setq dlam-cos (* cosp (- lam lam0)))

  (setq e4 (* e2 e2))
  (setq e6 (* e4 e2))

  ;; Arco meridiano (Snyder): M = a*(m1 - m2 + m3 - m4)
  (setq m1 (* (- 1.0 (/ e2 4.0) (/ (* 3.0 e4) 64.0) (/ (* 5.0 e6) 256.0)) phi))
  (setq m2 (* (+ (/ (* 3.0 e2) 8.0) (/ (* 3.0 e4) 32.0) (/ (* 45.0 e6) 1024.0)) (sin (* 2.0 phi))))
  (setq m3 (* (+ (/ (* 15.0 e4) 256.0) (/ (* 45.0 e6) 1024.0)) (sin (* 4.0 phi))))
  (setq m4 (* (/ (* 35.0 e6) 3072.0) (sin (* 6.0 phi))))
  (setq M (* semi-major (+ m1 (* -1.0 m2) m3 (* -1.0 m4))))

  (setq easting
    (+ 500000.0
       (* k0 N
          (+
            dlam-cos
            (/ (* (- 1.0 tan2 (- C)) (expt dlam-cos 3.0)) 6.0)
            (/ (* (+ 5.0 (* -18.0 tan2) (* tan2 tan2) (* 72.0 C) (* -58.0 ep2)) (expt dlam-cos 5.0)) 120.0)
          )
       )
    )
  )

  (setq northing
    (* k0
       (+
         M
         (* N tanp
            (+
              (/ (* dlam-cos dlam-cos) 2.0)
              (/ (* (+ 5.0 (- tan2) (* 9.0 C) (* 4.0 C C)) (expt dlam-cos 4.0)) 24.0)
              (/ (* (+ 61.0 (* -58.0 tan2) (* tan2 tan2) (* 600.0 C) (* -330.0 ep2)) (expt dlam-cos 6.0)) 720.0)
            )
         )
       )
    )
  )

  (if (= hemisphere "S")
    (setq northing (- 10000000.0 northing))
  )

  (setq epsg (itoa (+ zone (if (= hemisphere "S") 31960 31954))))
  (setq utm-code (strcat hemisphere (itoa zone)))

  (list zone hemisphere easting northing epsg utm-code)
)

(defun c:LatLongParaUTM_Ponto (/ raw pair lat lon result zone hemisphere easting northing epsg utm-code)
  (princ "\n\nIniciando c:LatLongParaUTM_Ponto...")

  (setq pair nil)
  (while (null pair)
    (setq raw (getstring T "\nDigite Latitude,Longitude (ex.: -25.580205, -52.978574): "))
    (setq pair (AGM:ParseLatLonPair raw))
    (if (null pair)
      (princ "\nFormato invalido. Use: latitude, longitude (graus decimais).")
    )
  )

  (setq lat (nth 0 pair))
  (setq lon (nth 1 pair))

  (if (or (< lat -90.0) (> lat 90.0) (< lon -180.0) (> lon 180.0))
    (progn
      (princ "\nValores fora da faixa valida. Latitude [-90,90], Longitude [-180,180].")
    )
    (progn
      (setq result (AGM:LatLonToUTM lat lon))
      (setq zone (nth 0 result))
      (setq hemisphere (nth 1 result))
      (setq easting (nth 2 result))
      (setq northing (nth 3 result))
      (setq epsg (nth 4 result))
      (setq utm-code (nth 5 result))

      (entmakex
        (list
          '(0 . "POINT")
          (cons 10 (list easting northing 0.0))
        )
      )

      (princ (strcat "\nUTM inferido: SIRGAS 2000 / UTM " utm-code " (EPSG:" epsg ")"))
      (princ (strcat "\nEasting:  " (rtos easting 2 3)))
      (princ (strcat "\nNorthing: " (rtos northing 2 3)))
      (princ "\nPonto criado com sucesso no desenho.")
    )
  )

  (princ "\nFinalizando c:LatLongParaUTM_Ponto.")
  (princ)
)
