package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.config.ShareProperties;
import com.xc.luckysheet.server.contract.ShareCreateRequest;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.store.ShareRow;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.junit.jupiter.api.Test;

import java.time.Duration;
import java.time.Instant;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.doAnswer;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class GuestShareServiceTest {
    @Test
    void tokenIsOpaqueToStorageAndResolvesOnlyToServerRole() {
        WorkbookStore store = mock(WorkbookStore.class);
        WorkbookLifecycleService lifecycle = mock(WorkbookLifecycleService.class);
        WorkbookAuthorizationService authorization = mock(WorkbookAuthorizationService.class);
        when(authorization.role("unit-1", "owner")).thenReturn(Optional.of(WorkbookAclRole.OWNER));
        ShareProperties properties = new ShareProperties(Duration.ofHours(1), Duration.ofDays(7));
        GuestShareService service = new GuestShareService(store, properties, lifecycle, authorization);
        ShareRow[] stored = new ShareRow[1];
        doAnswer(invocation -> { stored[0] = invocation.getArgument(0); return null; }).when(store).insertShare(any(ShareRow.class));
        var response = service.create("unit-1", new ShareCreateRequest("viewer", null), "owner");
        assertFalse(stored[0].tokenHash().contains(response.token()));
        when(store.findActiveShare(eq(response.shareId()), any(Instant.class))).thenReturn(Optional.of(stored[0]));
        when(store.findActiveShare(eq("unit-1"), eq(response.shareId()), any(Instant.class))).thenReturn(Optional.of(stored[0]));

        GuestShareService.GuestIdentity identity = service.authenticate(response.token());
        assertEquals("guest:" + response.shareId(), identity.subject());
        assertEquals(WorkbookAclRole.VIEWER, service.roleFor("unit-1", identity.subject()));
    }

    @Test
    void expiredOrMalformedTokenIsRejected() {
        WorkbookStore store = mock(WorkbookStore.class);
        GuestShareService service = new GuestShareService(store, new ShareProperties(Duration.ofHours(1), Duration.ofDays(7)),
                mock(WorkbookLifecycleService.class), mock(WorkbookAuthorizationService.class));
        assertThrows(ServiceException.class, () -> service.authenticate("not-a-token"));
    }
}
