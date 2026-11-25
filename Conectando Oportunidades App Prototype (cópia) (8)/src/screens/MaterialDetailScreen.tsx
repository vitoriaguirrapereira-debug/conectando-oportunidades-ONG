import type { ScreenProps } from "../App"
import { getGuide, opportunities } from "../data/opportunities"

const ONG_URL = "https://linktr.ee/ongnovamulher"

export default function MaterialDetailScreen({
  navigate,
  state,
  openOpportunity,
}: ScreenProps) {
  const guide = getGuide(state.selectedGuideId)
  const related = guide.courseIds
    .map((id) => opportunities.find((item) => item.id === id))
    .filter((item): item is (typeof opportunities)[number] => Boolean(item))
  return (
    <div className="material-detail">
      <header>
        <button type="button" onClick={() => navigate("materials")}>
          Voltar aos materiais
        </button>
        <span aria-hidden="true">{guide.emoji}</span>
        <small>
          {guide.category} · {guide.time}
        </small>
        <h1>{guide.title}</h1>
        <p>{guide.introduction}</p>
      </header>
      <main>
        <section className="guide-steps">
          <span className="eyebrow">
            {guide.stepLabel ?? "Checklist prático"}
          </span>
          <h2>{guide.stepTitle ?? "Faça no seu ritmo"}</h2>
          {guide.steps.map((step, index) => (
            <div className="guide-step" key={index}>
              <span>{index + 1}</span>
              <p>
                {typeof step === "string" ? (
                  step
                ) : (
                  <>
                    <strong>{step.title}:</strong> {step.text}
                  </>
                )}
              </p>
            </div>
          ))}
          {guide.tips && (
            <aside className="guide-tips">
              <span className="eyebrow">Dicas práticas</span>
              <h3>Para deixar o resultado ainda melhor</h3>
              {guide.tips.map((tip) => (
                <div key={tip.title}>
                  <strong>{tip.title}:</strong> {tip.text}
                </div>
              ))}
            </aside>
          )}
        </section>
        <section className="related-courses">
          <span className="eyebrow">Continue aprendendo</span>
          <h2>Cursos gratuitos relacionados</h2>
          {related.map((item) => (
            <button
              key={item.id}
              type="button"
              onClick={() => openOpportunity(item.id)}
            >
              <span aria-hidden="true">{item.emoji}</span>
              <span>
                <strong>{item.title}</strong>
                <small>{item.institution}</small>
              </span>
              <em>Ver curso</em>
            </button>
          ))}
        </section>
        <section className="ong-support-card">
          <span className="eyebrow">Apoio e conexão</span>
          <h2>Fale com a ONG Nova Mulher</h2>
          <p>
            Acesse as redes e os canais oficiais para conhecer iniciativas,
            atividades e formas de apoio.
          </p>
          <a href={ONG_URL} target="_blank" rel="noopener noreferrer">
            Abrir canais oficiais
          </a>
        </section>
      </main>
    </div>
  )
}
