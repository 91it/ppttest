Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById('search-button').addEventListener('click', async () => {
      const keyword = document.getElementById('search-keyword').value;
      const resultsDiv = document.getElementById('image-results');
      resultsDiv.innerHTML = 'Searching...';
      try {
        const images = await searchImagesFromBackend(keyword);
        resultsDiv.innerHTML = '';
        images.forEach(image => {
          const div = document.createElement('div');
          const img = document.createElement('img');
          const button = document.createElement('button');
          img.src = 'data:image/svg+xml;base64,' + btoa(image.data);
          img.style.width = '100px';
          img.style.margin = '10px';
          button.textContent = 'Insert';
          button.addEventListener('click', () => insertImage(img.src));
          div.appendChild(img);
          div.appendChild(button);
          resultsDiv.appendChild(div);
        });
      } catch (error) {
        resultsDiv.innerHTML = 'Error occurred while searching images.';
      }
    });
  }
});
async function searchImagesFromBackend(keyword) {
  const response = await fetch(`https://3af2893f61064f289daa302c5bf46dc0.apig.la-south-2.huaweicloudapis.com/search?keyword=${keyword}`);
  //const response = await fetch(`http://localhost:5000/search?keyword=${keyword}`);
  const images = await response.json();
  return images;
}
async function insertImage(svgDataUrl) {
  const img = new Image();
  img.src = svgDataUrl;
  img.onload = async () => {
    const canvas = document.createElement('canvas');
    canvas.width = img.width;
    canvas.height = img.height;
    const ctx = canvas.getContext('2d');
    if (ctx) {
      ctx.drawImage(img, 0, 0);
      const pngDataUrl = canvas.toDataURL('image/png');
      const base64Data = pngDataUrl.replace(/^data:image\/png;base64,/, "");
      Office.context.document.setSelectedDataAsync(base64Data, {
        coercionType: Office.CoercionType.Image
      }, asyncResult => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          setMessage("Error: " + asyncResult.error.message);
        }
      });
    }
  };
}