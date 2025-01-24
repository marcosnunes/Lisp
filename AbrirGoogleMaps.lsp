(defun c:AbrirGoogleMaps (/ pt coordenadas easting northing lat long url wshell proj4-script temp-html-file html-content file epsg-code)
  (princ "\n\nIniciando c:AbrirGoogleMaps...")
  (princ "\n")

  ;; Solicita ao usuario o sistema de coordenadas
  (initget "31982 31983")
  (setq epsg-code (getkword "\nEscolha o sistema de coordenadas (31982/31983) [31982]: "))
  (if (null epsg-code)
    (setq epsg-code "31982") ; Define como 31982 se o usuário nao escolher
  )
  
  (princ (strcat "\n\nSistema de coordenadas escolhido: EPSG:" epsg-code))
  (princ "\n")

  (setq pt (getpoint "\nSelecione um ponto para abrir no Google Maps: "))
  (if pt
    (progn
      (setq coordenadas (trans pt 1 0))
      (setq easting (float (car coordenadas)))
      (setq northing (float (cadr coordenadas)))

      (princ (strcat "\n  Easting:  " (rtos easting 2 3)))
      (princ (strcat "\n  Northing: " (rtos northing 2 3)))
      (princ "\n")

      ;; Inicia bloco de conversao de UTM para lat/long com Proj4js e gera o botao
      (setq proj4-script (strcat
        "<!DOCTYPE html>
        <html lang='pt-br'>
        <head>
          <meta charset='UTF-8'>
          <title>Localizacao no Google Maps</title>
          <meta http-equiv='Cache-Control' content='no-cache, no-store, must-revalidate'>
          <meta http-equiv='Pragma' content='no-cache'>
          <meta http-equiv='Expires' content='0'>
          <script src='https://cdnjs.cloudflare.com/ajax/libs/proj4js/2.7.5/proj4.js'></script>
          <style>
            body { font-family: sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background-color: #f0f0f0; }
            .container { text-align: center; padding: 20px; background-color: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
            #mapButton { padding: 10px 20px; background-color: #4CAF50; color: white; border: none; border-radius: 4px; cursor: pointer; margin-top: 20px; }
            #mapButton:hover { background-color: #45a049; }
            #coordinates { margin-top: 10px; font-size: 1.2em; }
          </style>
        </head>
        <body>
            <div class='container'>
              <h1>Localizacao no Google Maps</h1>
              <div id='coordinates'>
                <p>Aguarde... Calculando coordenadas.</p>
              </div>
              <button id='mapButton' disabled>Abrir no Google Maps</button>
            </div>
          <script>
             function convertUTMtoLatLon(easting, northing, epsgCode) {
                var proj4String = '';
               if(epsgCode === '31982'){
                  proj4String = '+proj=utm +zone=22 +south +ellps=GRS80 +towgs84=0,0,0,0,0,0,0 +units=m +no_defs';
               } else if (epsgCode === '31983') {
                 proj4String = '+proj=utm +zone=23 +south +ellps=GRS80 +towgs84=0,0,0,0,0,0,0 +units=m +no_defs';
                }

              proj4.defs('EPSG:' + epsgCode, proj4String);
              var sourceCoords = [easting, northing];
              var destCoords = proj4('EPSG:' + epsgCode, 'EPSG:4326', sourceCoords);
              
              var latitude = destCoords[1];
              var longitude = destCoords[0];

              document.getElementById('coordinates').innerHTML = '<p>Latitude: ' + latitude.toFixed(6) + '<br>Longitude: ' + longitude.toFixed(6) + '</p>';

              var mapButton = document.getElementById('mapButton');
              mapButton.disabled = false;
              mapButton.textContent = 'Abrir no Google Maps';

              // Formata a URL corretamente com coordenadas decimais
              
                var latDeg = Math.abs(latitude);
                var latMin = (latDeg - Math.floor(latDeg)) * 60;
                var latSec = (latMin - Math.floor(latMin)) * 60;
                var latDir = latitude >= 0 ? 'N' : 'S';
                var longDeg = Math.abs(longitude);
                var longMin = (longDeg - Math.floor(longDeg)) * 60;
                var longSec = (longMin - Math.floor(longMin)) * 60;
                var longDir = longitude >= 0 ? 'E' : 'W';
             
                var latString = Math.floor(latDeg) + '°' + Math.floor(latMin) + \"'\" + latSec.toFixed(1) + '\"' + latDir;
                var longString = Math.floor(longDeg) + '°' + Math.floor(longMin) + \"'\" + longSec.toFixed(1) + '\"' + longDir;
                

                var url = 'https://www.google.com/maps/place/' + encodeURIComponent(latString) + '+' + encodeURIComponent(longString) + '/@' + latitude.toFixed(8) + ',' + longitude.toFixed(8) + ',17z';
                 

              mapButton.onclick = function() {
                 window.open(url); // Remove '_blank' para abrir na mesma janela
              };
            }

            window.onload = function() {
              convertUTMtoLatLon(" (rtos easting 2 10) ", " (rtos northing 2 10) ",  \"" epsg-code "\");
            }
          </script>
        </body>
        </html>"
      ))

      ;; Cria o arquivo HTML temporário
      (setq temp-html-file (vl-filename-mktemp "temp_proj4.html"))

      ;; Escreve o HTML no arquivo
      (setq file (open temp-html-file "w"))
      (if file
        (progn
          (write-line proj4-script file)
          (close file)
        )
        (progn
          (princ (strcat "\n  Erro ao criar/escrever no arquivo HTML: " temp-html-file))
          (exit) ; Aborta se nao conseguir criar o arquivo
        )
      )
      (princ "\n")

      ;; Abre o navegador padrao usando o comando BROWSER
      (command "_.BROWSER" temp-html-file) ; Abre o arquivo HTML no navegador
      
      (princ "\n\nHTML gerado e aberto no navegador. Clique no botao para abrir o Google Maps.")
      (princ "\n")
    )
    (princ "\nNenhum ponto selecionado.")
  )
  (princ "\n\nFinalizando c:AbrirGoogleMaps.")
  (princ "\nCommand: ")
  (princ)
)
