import type { ScreenProps } from "../App"
import { guides } from "../data/opportunities"
import { useState } from "react"
import {
  downloadOrdersForm,
  downloadResumeTemplate,
  downloadSalesWorkbook,
} from "../utils/financialDownloads"

export default function MaterialsScreen({ openGuide }: ScreenProps) {
  const [downloadMessage, setDownloadMessage] = useState("")
  const downloads = [
    {
      title: "Planilha de Controle de Vendas",
      description: "Registre suas vendas e acompanhe o que já recebeu. Duas abas: Vendas e Resumo, com totais e fórmulas automáticas por forma de pagamento.",
      format: "Excel · .xlsx",
      buttonLabel: "Baixar planilha Excel",
      action: downloadSalesWorkbook,
    },
    {
      title: "Ficha de Controle de Pedidos e Parcelas",
      description: "Organize dados da cliente, produtos, pagamentos e até quatro parcelas. Abra o arquivo baixado para imprimir em A4 ou salvar como PDF.",
      format: "HTML · impressão A4",
      buttonLabel: "Baixar ficha para impressão",
      action: downloadOrdersForm,
    },
    {
      title: "Modelo de Currículo",
      description:
        "Um modelo organizado com exemplos e dicas para ajudar você a montar seu currículo. É só baixar, clicar nos textos para preencher seus dados e salvar ou imprimir.",
      format: "MODELO PRONTO · A4",
      buttonLabel: "Acessar modelo de currículo",
      action: downloadResumeTemplate,
    },
  ]
  return (
    <div className="materials-screen page-surface">
      <header className="page-heading">
        <div>
          <span className="eyebrow">Ferramentas para o dia a dia</span>
          <h1>Guias Práticos de Apoio ao Negócio</h1>
          <p>
            Orientações objetivas para colocar em prática hoje e aprofundar com
            cursos gratuitos da curadoria.
          </p>
        </div>
      </header>
      <section aria-labelledby="downloads-title" className="px-6 pt-6 lg:px-12">
        <h2 id="downloads-title" className="font-heading text-2xl font-semibold text-slate-900">Planilhas e Materiais para Baixar</h2>
        <p className="mt-2 text-sm leading-relaxed text-slate-600">Grátis, sem cadastro e gerados no seu navegador. Os dados preenchidos nos arquivos ficam com você, não são enviados a plataforma.</p>
        <div className="mt-4 grid items-stretch gap-4 md:grid-cols-2 xl:grid-cols-3">
          {downloads.map((item) => (
            <article key={item.title} className="flex h-full min-w-0 flex-col rounded-2xl border border-teal-200 bg-white p-5">
              <span className="text-xs font-bold uppercase tracking-wide text-teal-800">{item.format}</span>
              <h3 className="mt-2 font-heading text-xl font-semibold text-slate-900">{item.title}</h3>
              <p className="mt-2 mb-5 flex-1 text-sm leading-relaxed text-slate-600">{item.description}</p>
              <button type="button" className="min-h-11 rounded-xl bg-teal-800 px-4 py-3 text-sm font-semibold text-white focus-visible:ring-2"
                onClick={() => {
                  try {
                    item.action()
                    setDownloadMessage(`Download solicitado: ${item.title}. Confira os downloads do navegador.`)
                  } catch {
                    setDownloadMessage("Não foi possível gerar o arquivo. Tente novamente neste navegador.")
                  }
                }}>
                {item.buttonLabel}
              </button>
            </article>
          ))}
        </div>
        <p role="status" className="mt-3 text-sm text-teal-900">{downloadMessage}</p>
        <p className="mt-2 text-sm text-slate-600">Na planilha, apague a venda de exemplo antes de começar e mantenha as fórmulas da aba Resumo. Use um aplicativo compatível com Excel para editar.</p>
      </section>
      <section className="guide-grid" data-tutorial="materials-featured">
        {guides.map((guide) => (
          <button
            key={guide.id}
            type="button"
            onClick={() => openGuide(guide.id)}
            style={{ "--guide-color": guide.color } as React.CSSProperties}
          >
            <span className="guide-grid__icon" aria-hidden="true">
              {guide.emoji}
            </span>
            <span className="eyebrow">{guide.category}</span>
            <strong>{guide.title}</strong>
            <p>{guide.cardDescription ?? guide.introduction}</p>
            <small>{guide.time} de leitura · Abrir guia</small>
          </button>
        ))}
      </section>
    </div>
  )
}
