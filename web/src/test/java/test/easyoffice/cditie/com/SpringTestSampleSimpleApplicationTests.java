package test.easyoffice.cditie.com;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.context.ApplicationContext;
import org.springframework.test.context.junit4.SpringRunner;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * Created by zhuyunhui on 9/13/2017.
 */
@RunWith(SpringRunner.class)
@SpringBootTest
public class SpringTestSampleSimpleApplicationTests {
	@Autowired
	ApplicationContext ctx;

	@Test
	public void testContextLoads() throws Exception {
		assertThat(this.ctx).isNotNull();
		assertThat(this.ctx.containsBean("helloWorldService")).isTrue();
		assertThat(this.ctx.containsBean("sampleSimpleApplication")).isTrue();
	}
}
