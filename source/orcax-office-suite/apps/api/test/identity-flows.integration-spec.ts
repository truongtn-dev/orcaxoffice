import 'reflect-metadata';
import assert from 'node:assert/strict';
import { after, before, describe, it } from 'node:test';
import { INestApplication } from '@nestjs/common';
import { Test } from '@nestjs/testing';
import { randomUUID } from 'crypto';
import request from 'supertest';
import { DataSource } from 'typeorm';
import { AppModule } from '../src/app.module';
import { configureHttp } from '../src/bootstrap/configure-http';
import { CORRELATION_ID_HEADER } from '@orcax/api-contracts';

describe('identity workspace flows', () => {
  let app: INestApplication;
  let dataSource: DataSource;
  const password = 'correct-horse-battery';
  const ownerEmail = `owner-${randomUUID()}@example.com`;
  const memberEmail = `member-${randomUUID()}@example.com`;

  before(
    async () => {
      const moduleRef = await Test.createTestingModule({
        imports: [AppModule],
      }).compile();
      app = moduleRef.createNestApplication({ logger: false });
      configureHttp(app, { webOrigin: 'http://localhost:3000' });
      await app.init();
      dataSource = app.get(DataSource);
    },
    { timeout: 60_000 },
  );

  after(async () => {
    await app?.close();
  });

  it(
    'registers, verifies, switches with confirmation, and keeps workspaces isolated',
    { timeout: 120_000 },
    async () => {
      const agent = request.agent(app.getHttpServer());

      const register = await agent
        .post('/api/v1/auth/register')
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({
          email: ownerEmail,
          password,
          displayName: 'Chủ sở hữu',
        })
        .expect(200);
      assert.equal(typeof register.body.data.message, 'string');

      const resendUnknown = await agent
        .post('/api/v1/auth/email/resend')
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ email: `missing-${randomUUID()}@example.com` })
        .expect(200);
      assert.equal(resendUnknown.body.data.message, register.body.data.message);

      const verifyCode = await readCode(agent, ownerEmail, 'Mã xác minh');
      await agent
        .post('/api/v1/auth/email/verify')
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ email: ownerEmail, code: verifyCode })
        .expect(200);

      const badLogin = await agent
        .post('/api/v1/auth/login')
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ email: ownerEmail, password: 'wrong-password-xx' })
        .expect(401);
      assert.equal(badLogin.body.error.code, 'UNAUTHORIZED');

      const login = await agent
        .post('/api/v1/auth/login')
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ email: ownerEmail, password })
        .expect(200);
      let accessToken = login.body.data.accessToken as string;
      const personalId = login.body.data.workspace.id as string;
      assert.equal(login.body.data.workspace.kind, 'personal');

      const me = await agent
        .get('/api/v1/auth/me')
        .set('authorization', `Bearer ${accessToken}`)
        .expect(200);
      assert.equal(me.body.data.workspace.id, personalId);

      const created = await agent
        .post('/api/v1/organizations')
        .set('authorization', `Bearer ${accessToken}`)
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ name: 'Công ty OrcaX' })
        .expect(201);
      const organizationWorkspaceId = created.body.data.id as string;
      assert.equal(created.body.data.kind, 'organization');

      const blockedSwitch = await agent
        .post('/api/v1/workspaces/switch')
        .set('authorization', `Bearer ${accessToken}`)
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ workspaceId: organizationWorkspaceId, confirmed: false })
        .expect(400);
      assert.equal(blockedSwitch.body.error.code, 'WORKSPACE_CONFIRMATION_REQUIRED');

      const previousToken = accessToken;
      const switched = await agent
        .post('/api/v1/workspaces/switch')
        .set('authorization', `Bearer ${accessToken}`)
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ workspaceId: organizationWorkspaceId, confirmed: true })
        .expect(200);
      accessToken = switched.body.data.accessToken as string;
      assert.equal(switched.body.data.workspace.id, organizationWorkspaceId);

      await agent
        .get('/api/v1/auth/me')
        .set('authorization', `Bearer ${previousToken}`)
        .expect(401);

      await agent
        .post('/api/v1/organizations')
        .set('authorization', `Bearer ${accessToken}`)
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ name: 'Không được tạo từ org' })
        .expect(403);

      await agent
        .post('/api/v1/invitations')
        .set('authorization', `Bearer ${accessToken}`)
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ email: memberEmail, role: 'member' })
        .expect(201);

      const inviteToken = await readInviteToken(agent, memberEmail);

      await registerAndVerify(agent, memberEmail, 'Thành viên');
      const memberLogin = await agent
        .post('/api/v1/auth/login')
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ email: memberEmail, password })
        .expect(200);
      let memberToken = memberLogin.body.data.accessToken as string;

      const accepted = await agent
        .post('/api/v1/invitations/accept')
        .set('authorization', `Bearer ${memberToken}`)
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ token: inviteToken })
        .expect(200);
      assert.equal(accepted.body.data.workspaceId, organizationWorkspaceId);

      const memberSwitch = await agent
        .post('/api/v1/workspaces/switch')
        .set('authorization', `Bearer ${memberToken}`)
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ workspaceId: organizationWorkspaceId, confirmed: true })
        .expect(200);
      memberToken = memberSwitch.body.data.accessToken as string;

      const members = await agent
        .get('/api/v1/members')
        .set('authorization', `Bearer ${accessToken}`)
        .expect(200);
      const memberRow = (members.body.data.members as Array<{ id: string; email: string }>).find(
        (row) => row.email === memberEmail,
      );
      assert.ok(memberRow);

      await agent
        .post(`/api/v1/memberships/${memberRow.id}/role`)
        .set('authorization', `Bearer ${accessToken}`)
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ role: 'auditor' })
        .expect(200);

      await agent
        .get('/api/v1/members')
        .set('authorization', `Bearer ${memberLogin.body.data.accessToken}`)
        .expect(401);

      const memberPersonalLogin = await agent
        .post('/api/v1/auth/login')
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ email: memberEmail, password })
        .expect(200);
      assert.equal(memberPersonalLogin.body.data.workspace.kind, 'personal');
      const personalMembers = await agent
        .get('/api/v1/members')
        .set('authorization', `Bearer ${memberPersonalLogin.body.data.accessToken}`)
        .expect(200);
      assert.equal(
        (personalMembers.body.data.members as Array<{ email: string }>).some(
          (row) => row.email === ownerEmail,
        ),
        false,
      );
      assert.equal(
        (personalMembers.body.data.members as Array<{ email: string }>).some(
          (row) => row.email === memberEmail,
        ),
        true,
      );

      const sessions = await agent
        .get('/api/v1/sessions')
        .set('authorization', `Bearer ${accessToken}`)
        .expect(200);
      const current = (sessions.body.data.sessions as Array<{ id: string; current: boolean }>).find(
        (row) => row.current,
      );
      assert.ok(current);

      await agent
        .delete(`/api/v1/sessions/${current.id}`)
        .set('authorization', `Bearer ${accessToken}`)
        .set(CORRELATION_ID_HEADER, randomUUID())
        .expect(200);

      await agent.get('/api/v1/auth/me').set('authorization', `Bearer ${accessToken}`).expect(401);

      const audits = await dataSource.query(
        `SELECT action, outcome FROM audit_events
         WHERE action = ANY($1)
         ORDER BY occurred_at ASC`,
        [
          [
            'identity.registered',
            'identity.login.failed',
            'identity.login.succeeded',
            'workspace.created',
            'workspace.switched',
            'membership.invited',
            'membership.accepted',
            'membership.role.changed',
            'session.revoked',
          ],
        ],
      );
      const actions = new Set((audits as Array<{ action: string }>).map((row) => row.action));
      for (const required of [
        'identity.login.failed',
        'identity.login.succeeded',
        'workspace.switched',
        'membership.role.changed',
        'session.revoked',
      ]) {
        assert.equal(actions.has(required), true, `missing audit ${required}`);
      }
    },
  );

  it(
    'resets a password with a generic forgot response and revokes sessions',
    { timeout: 60_000 },
    async () => {
      const agent = request.agent(app.getHttpServer());
      const email = `reset-${randomUUID()}@example.com`;
      await registerAndVerify(agent, email, 'Đặt lại');
      const login = await agent
        .post('/api/v1/auth/login')
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ email, password })
        .expect(200);
      const token = login.body.data.accessToken as string;

      const forgotKnown = await agent
        .post('/api/v1/auth/password/forgot')
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ email })
        .expect(200);
      const forgotUnknown = await agent
        .post('/api/v1/auth/password/forgot')
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ email: `ghost-${randomUUID()}@example.com` })
        .expect(200);
      assert.equal(forgotKnown.body.data.message, forgotUnknown.body.data.message);

      const code = await readCode(agent, email, 'Mã đặt lại mật khẩu');
      const nextPassword = 'replacement-password';
      await agent
        .post('/api/v1/auth/password/reset')
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ email, code, password: nextPassword })
        .expect(200);

      await agent.get('/api/v1/auth/me').set('authorization', `Bearer ${token}`).expect(401);
      await agent
        .post('/api/v1/auth/login')
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ email, password })
        .expect(401);
      await agent
        .post('/api/v1/auth/login')
        .set(CORRELATION_ID_HEADER, randomUUID())
        .send({ email, password: nextPassword })
        .expect(200);
    },
  );
});

async function registerAndVerify(
  agent: ReturnType<typeof request.agent>,
  email: string,
  displayName: string,
): Promise<void> {
  await agent
    .post('/api/v1/auth/register')
    .set(CORRELATION_ID_HEADER, randomUUID())
    .send({ email, password: 'correct-horse-battery', displayName })
    .expect(200);
  const code = await readCode(agent, email, 'Mã xác minh');
  await agent
    .post('/api/v1/auth/email/verify')
    .set(CORRELATION_ID_HEADER, randomUUID())
    .send({ email, code })
    .expect(200);
}

async function readCode(
  agent: ReturnType<typeof request.agent>,
  email: string,
  marker: string,
): Promise<string> {
  const mailbox = await agent.get('/api/v1/dev/mailbox').query({ email }).expect(200);
  const messages = mailbox.body.data.messages as Array<{ text: string }>;
  const latest = [...messages].reverse().find((item) => item.text.includes(marker));
  assert.ok(latest, `mailbox missing ${marker}`);
  const match = latest.text.match(/\b(\d{6})\b/);
  assert.ok(match, 'verification code missing');
  return match[1];
}

async function readInviteToken(
  agent: ReturnType<typeof request.agent>,
  email: string,
): Promise<string> {
  const mailbox = await agent.get('/api/v1/dev/mailbox').query({ email }).expect(200);
  const messages = mailbox.body.data.messages as Array<{ text: string }>;
  const latest = [...messages].reverse().find((item) => item.text.includes('Mã lời mời'));
  assert.ok(latest, 'invite mail missing');
  const match = latest.text.match(/Mã lời mời:\s*([A-Za-z0-9_-]+)/);
  assert.ok(match, 'invite token missing');
  return match[1];
}
