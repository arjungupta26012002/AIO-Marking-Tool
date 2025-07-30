function doGet() {

  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Excelerate Marking System');
}
