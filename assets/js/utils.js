// Utilidades compartidas
async function readFileAsUint8Array(file) {
  if (!file) throw new Error('Archivo no válido');
  if (file.arrayBuffer) {
    const buf = await file.arrayBuffer();
    return new Uint8Array(buf);
  }
  return await new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(new Uint8Array(fr.result));
    fr.onerror = reject;
    fr.readAsArrayBuffer(file);
  });
}

function on(id, evt, fn) {
  const el = document.getElementById(id);
  if (!el) {
    console.warn(`⚠️ Elemento ${id} no encontrado para evento ${evt}`);
    return null;
  }
  el.addEventListener(evt, fn);
  return el;
}

function decodeText(bytes) {
  if (!bytes) return '';
  try {
    return new TextDecoder('utf-8').decode(bytes);
  } catch (err) {
    console.warn('TextDecoder no disponible, usando fallback ASCII');
    let result = '';
    for (let i = 0; i < bytes.length; i++) {
      result += String.fromCharCode(bytes[i]);
    }
    return result;
  }
}
