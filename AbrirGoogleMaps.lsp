(defun c:AbrirGoogleMaps (/ pt coordenadas easting northing proj4-script temp-html-file file)
  (princ "\n\nIniciando c:AbrirGoogleMaps...")
  (princ "\n")

  (princ "\nModo automatico: inferencia de zona SIRGAS 2000 / UTM (Brasil).")
  (princ "\n")

  (setq pt (getpoint "\nSelecione um ponto para abrir no Google Maps: "))
  (if pt
    (progn
      (setq coordenadas (trans pt 1 0))
      (setq easting (float (car coordenadas)))
      (setq northing (float (cadr coordenadas)))

      ;; Faixas UTM tipicas para reduzir abertura de pontos inconsistentes.
      (if (or (< easting 100000.0) (> easting 900000.0) (< northing 0.0) (> northing 10000000.0))
        (progn
          (princ "\nCoordenadas fora da faixa UTM esperada (E: 100000-900000, N: 0-10000000).")
          (princ "\nVerifique UCS/CRS antes de continuar.")
          (setq pt nil)
        )
      )

      (if pt
        (progn
          (princ (strcat "\n  Easting:  " (rtos easting 2 3)))
          (princ (strcat "\n  Northing: " (rtos northing 2 3)))
          (princ "\n")

          ;; Inicia bloco de conversao de UTM para lat/long com Proj4js e gera o botao
          (setq proj4-script (strcat
            "<!DOCTYPE html>
        <html lang='pt-br'>
        <head>
          <meta charset='UTF-8'>
          <title>Localizaçao no Google Maps</title>
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
              <h1>Localizaçao no Google Maps</h1>
              <div id='coordinates'>
                <p>Aguarde... Calculando coordenadas.</p>
              </div>
              <button id='mapButton' disabled>Abrir no Google Maps</button>
            </div>
          <script>
             function getEpsgCode(hemisphere, zone) {
              return hemisphere === 'S' ? (31960 + zone) : (31954 + zone);
            }

            function inferBestZone(easting, northing) {
              // Heuristica para Brasil sem georreferencia: escolhe hemisferio por Northing
              // e avalia zonas 18..25, priorizando continuidade e fallback em S22.
              var hemisphere = northing < 1200000 ? 'N' : 'S';
              var zones = [18, 19, 20, 21, 22, 23, 24, 25];
              var brazilBBox = { minLat: -34.0, maxLat: 6.0, minLon: -74.5, maxLon: -32.0 };
              var candidates = [];

              for (var i = 0; i < zones.length; i++) {
                var zone = zones[i];
                var proj4String = '+proj=utm +zone=' + zone + (hemisphere === 'S' ? ' +south' : '') + ' +ellps=GRS80 +units=m +no_defs';
                var sourceCrs = 'SIRGAS2000_UTM_' + hemisphere + zone;
                proj4.defs(sourceCrs, proj4String);

                var dest = proj4(sourceCrs, 'EPSG:4326', [easting, northing]);
                var lat = dest[1];
                var lon = dest[0];
                var insideBrazil = (lat >= brazilBBox.minLat && lat <= brazilBBox.maxLat && lon >= brazilBBox.minLon && lon <= brazilBBox.maxLon);

                if (insideBrazil) {
                  candidates.push({ zone: zone, hemisphere: hemisphere, lat: lat, lon: lon });
                }
              }

              // Persistencia por sessao do navegador para manter coerencia em pontos sequenciais.
              var lastZone = parseInt(localStorage.getItem('agm_last_zone') || '22', 10);
              if (isNaN(lastZone) || lastZone < 18 || lastZone > 25) {
                lastZone = 22;
              }

              var best = null;
              if (candidates.length > 0) {
                var bestDist = 999;
                for (var j = 0; j < candidates.length; j++) {
                  var d = Math.abs(candidates[j].zone - lastZone);
                  if (d < bestDist) {
                    bestDist = d;
                    best = candidates[j];
                  }
                }
              }

              // Fallback seguro para S22 quando nao houver candidato no bbox do Brasil.
              if (!best) {
                var fallbackZone = 22;
                var fallbackProj = '+proj=utm +zone=' + fallbackZone + (hemisphere === 'S' ? ' +south' : '') + ' +ellps=GRS80 +units=m +no_defs';
                var fallbackCrs = 'SIRGAS2000_UTM_' + hemisphere + fallbackZone;
                proj4.defs(fallbackCrs, fallbackProj);
                var fallbackDest = proj4(fallbackCrs, 'EPSG:4326', [easting, northing]);
                best = { zone: fallbackZone, hemisphere: hemisphere, lat: fallbackDest[1], lon: fallbackDest[0] };
              }

              localStorage.setItem('agm_last_zone', String(best.zone));
              best.candidateCount = candidates.length;
              return best;
            }

            function convertUTMtoLatLon(easting, northing) {
              if (easting < 100000 || easting > 900000 || northing < 0 || northing > 10000000) {
                document.getElementById('coordinates').innerHTML = '<p>Coordenadas UTM fora da faixa esperada.</p>';
                return;
              }

              var best = inferBestZone(easting, northing);
              var zone = best.zone;
              var hemisphere = best.hemisphere;
              var utmCode = hemisphere + zone;
              var epsgCode = getEpsgCode(hemisphere, zone);
              var latitude = best.lat;
              var longitude = best.lon;
              
              var heuristicNote = best.candidateCount > 1
                ? '<br><small>Zona inferida por heuristica (' + best.candidateCount + ' candidatas no Brasil).</small>'
                : '<br><small>Zona inferida automaticamente.</small>';

              document.getElementById('coordinates').innerHTML = '<p>Sistema inferido: SIRGAS 2000 / UTM ' + utmCode + ' (EPSG:' + epsgCode + ')<br>Latitude: ' + latitude.toFixed(6) + '<br>Longitude: ' + longitude.toFixed(6) + heuristicNote + '</p>';

              var mapButton = document.getElementById('mapButton');
              mapButton.disabled = false;
              mapButton.textContent = 'Abrir no Google Maps';

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
                window.open(url);
              };
            }

            window.onload = function() {
              convertUTMtoLatLon(" (rtos easting 2 10) ", " (rtos northing 2 10) ");
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
              (exit)
            )
          )
          (princ "\n")

          ;; Abre o navegador padrao usando o comando BROWSER
          (command "_.BROWSER" temp-html-file)

          (princ "\n\nHTML gerado e aberto no navegador. Clique no botao para abrir o Google Maps.")
          (princ "\n")
        )
      )
    )
    (princ "\nNenhum ponto selecionado.")
  )
  (princ "\n\nFinalizando c:AbrirGoogleMaps.")
  (princ "\nCommand: ")
  (princ)
)