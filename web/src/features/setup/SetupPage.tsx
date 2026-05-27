import { Page, PageHero, StepCard } from '../../components/PageLayout';
import { BlueprintPage } from './BlueprintPage';
import { AgentPage } from './AgentPage';
import { AccessPackagePage } from './AccessPackagePage';
import { LcwPage } from './LcwPage';
import { ConditionalAccessPage } from './ConditionalAccessPage';
import { SummaryPage } from './SummaryPage';

export function SetupPage() {
  return (
    <Page>
      <PageHero
        title="Setup"
        subtitle="Run each step below in order. Each step provisions one of the resources you'll use across the Manage, Govern, and Protect journeys."
      />
      <StepCard number={1} label="Agent Blueprint"><BlueprintPage /></StepCard>
      <StepCard number={2} label="Agent ID"><AgentPage /></StepCard>
      <StepCard number={3} label="Catalog + Access Package"><AccessPackagePage /></StepCard>
      <StepCard number={4} label="Lifecycle Workflow"><LcwPage /></StepCard>
      <StepCard number={5} label="CSA + Conditional Access policy"><ConditionalAccessPage /></StepCard>
      <StepCard number={6} label="Summary"><SummaryPage /></StepCard>
    </Page>
  );
}

