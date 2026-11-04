(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  let source = null, name = '', rotation = 0, result = null, revision = 0, busy = false, seal = null;
  function controls() {
    for (const id of ['rotateLeft','rotateRight','resetRotation','generate','download']) $(id).disabled = !source || busy;
    $('clearImage').disabled = !source && !busy;
    $('previewSection').hidden = !source;
    $('fileInfo').textContent = source ? `${name}｜${source.naturalWidth} × ${source.naturalHeight} 像素｜旋转 ${rotation}°` : '未选择图片';
  }
  function loadImage(url) {
    return new Promise((resolve, reject) => {
      const image = new Image(), timer = setTimeout(() => finish(new Error('图片读取超时，请重新选择')), 15000);
      function finish(error) { clearTimeout(timer); image.onload = image.onerror = null; error ? reject(error) : resolve(image); }
      image.onload = () => finish(image.naturalWidth && image.naturalHeight ? null : new Error('图片没有有效尺寸'));
      image.onerror = () => finish(new Error('图片无法读取，请选择有效的 PNG、JPG、WebP 或 BMP 图片'));
      image.src = url;
    });
  }
  function reportCanvas() {
    const canvas = document.createElement('canvas'), swap = rotation === 90 || rotation === 270;
    canvas.width = swap ? source.naturalHeight : source.naturalWidth;
    canvas.height = swap ? source.naturalWidth : source.naturalHeight;
    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('浏览器无法绘制图片');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.translate(canvas.width / 2, canvas.height / 2); ctx.rotate(rotation * Math.PI / 180);
    ctx.drawImage(source, -source.naturalWidth / 2, -source.naturalHeight / 2);
    return canvas;
  }
  function show(canvas) {
    const preview = $('preview'); preview.width = canvas.width; preview.height = canvas.height;
    preview.getContext('2d').drawImage(canvas, 0, 0);
  }
  async function upload(file) {
    if (!file) return;
    const token = ++revision; let url;
    source = null; result = null; name = ''; rotation = 0; busy = true; controls();
    $('status').textContent = '正在读取图片…';
    try {
      if (!/\.(?:png|jpe?g|webp|bmp)$/i.test(file.name)) throw new Error('请选择 PNG、JPG、WebP 或 BMP 图片');
      url = URL.createObjectURL(file);
      const image = await loadImage(url);
      if (token !== revision) return;
      source = image; name = file.name; show(reportCanvas());
      $('status').textContent = '图片已加载，请检查方向后加章或一键下载。';
    } catch (error) {
      if (token === revision) { source = null; $('status').textContent = '读取失败：' + error.message; }
    } finally {
      if (url) URL.revokeObjectURL(url);
      if (token === revision) { busy = false; controls(); }
    }
  }
  function rotate(angle) {
    if (!source || busy) return;
    rotation = ((angle % 360) + 360) % 360; result = null;
    try { show(reportCanvas()); $('status').textContent = '方向已更新，请重新加章或一键下载。'; }
    catch (error) { $('preview').width = $('preview').height = 1; $('status').textContent = '旋转失败：' + error.message; }
    controls();
  }
  function download(blob) {
    const url = URL.createObjectURL(blob), link = document.createElement('a');
    link.href = url; link.download = name.replace(/\.[^.]+$/, '') + '_旋正已盖章.jpg';
    document.body.append(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(url), 1500);
  }
  async function generate(save) {
    if (!source || busy) return;
    if (save && result) { download(result); return; }
    const token = revision;
    busy = true; result = null; controls(); $('status').textContent = '正在生成盖章图片…';
    try {
      const canvas = reportCanvas();
      seal ||= await loadImage('assets/quality-seal.png');
      if (token !== revision) return;
      // Preserve the original tool's seal size, tilt and bottom-right placement.
      const width = Math.round(canvas.width * .18), height = Math.round(seal.naturalHeight * width / seal.naturalWidth);
      const extra = Math.ceil(Math.max(width, height) * .18), stamp = document.createElement('canvas');
      stamp.width = width + extra * 2; stamp.height = height + extra * 2;
      const ctx = stamp.getContext('2d');
      ctx.translate(stamp.width / 2, stamp.height / 2); ctx.rotate(-5 * Math.PI / 180);
      ctx.drawImage(seal, -width / 2, -height / 2, width, height);
      const output = canvas.getContext('2d'); output.setTransform(1, 0, 0, 1, 0, 0);
      output.drawImage(stamp, canvas.width - stamp.width - Math.round(canvas.width * .055), canvas.height - stamp.height - Math.round(canvas.height * .045));
      const blob = await new Promise((resolve, reject) => canvas.toBlob(value => value ? resolve(value) : reject(new Error('图片生成失败，请重试')), 'image/jpeg', .95));
      if (token !== revision) return;
      result = blob; show(canvas); $('status').textContent = '已生成盖章图片，请检查预览。';
      if (save) download(blob);
    } catch (error) {
      if (token === revision) {
        result = null;
        try { show(reportCanvas()); } catch (_) { $('preview').width = $('preview').height = 1; }
        $('status').textContent = '加章失败：' + error.message;
      }
    } finally {
      if (token === revision) { busy = false; controls(); }
    }
  }
  $('reportFile').addEventListener('change', event => { upload(event.target.files[0]); event.target.value = ''; });
  const zone = $('dropZone');
  zone.addEventListener('dragover', event => { event.preventDefault(); zone.classList.add('drag'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag'));
  zone.addEventListener('drop', event => { event.preventDefault(); zone.classList.remove('drag'); upload(event.dataTransfer.files[0]); });
  $('rotateLeft').addEventListener('click', () => rotate(rotation - 90));
  $('rotateRight').addEventListener('click', () => rotate(rotation + 90));
  $('resetRotation').addEventListener('click', () => rotate(0));
  $('clearImage').addEventListener('click', () => {
    revision++; source = result = null; name = ''; rotation = 0; busy = false;
    $('preview').width = $('preview').height = 1; $('reportFile').value = ''; controls();
    $('status').textContent = '图片已清空，请重新上传。';
  });
  $('generate').addEventListener('click', () => generate(false));
  $('download').addEventListener('click', () => generate(true));
  controls();
  $('reportFile').disabled = false;
  $('status').textContent = '请上传检验报告单图片。';
})();
