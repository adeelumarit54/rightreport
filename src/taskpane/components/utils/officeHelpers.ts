export async function insertIntoWord(inspection: any) {
  await Word.run(async (context) => {
    const placeholders = [
      { key: "{{client}}", value: inspection.client },
      { key: "{{site}}", value: inspection.site },
      { key: "{{date}}", value: inspection.date },
      { key: "{{findings}}", value: inspection.findings },
    ];

    for (const p of placeholders) {
      const results = context.document.body.search(p.key, { matchCase: false });
      results.load("items");
      await context.sync();

      results.items.forEach((item) => {
        item.insertText(p.value || "", "Replace");
      });
    }

    await context.sync();
  });
}
